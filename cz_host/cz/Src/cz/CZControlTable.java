package cz;

import java.awt.Component;
import java.awt.Rectangle;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.awt.event.WindowAdapter;
import java.awt.event.WindowEvent;
import java.text.DecimalFormat;
import java.util.Locale;
import java.util.Vector;

import javax.swing.JButton;
import javax.swing.JComboBox;
import javax.swing.JFrame;
import javax.swing.JDialog;
import javax.swing.JLabel;
import javax.swing.JOptionPane;
import javax.swing.JScrollPane;
// add start 2008.09.12
import javax.swing.JScrollBar;
// add end 2008.09.12
import javax.swing.JTable;
import javax.swing.JTextField;
import javax.swing.ListSelectionModel;
import javax.swing.Timer;
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

import czclass.CZNativeHikiage;
import czclass.CZParamControlDefine;
import czclass.CZParamControlT6Define;
import czclass.CZParamT6Table;
import czclass.CZResult;

/**
 *   ����e�[�u���ύXWindow
 * @author  (KPK Co.,Ltd.)
 * @version 1.0 (2003/04/01)
 * 2008.09.10 H.Nagamine �e�[�u���I����ʐ���
 * 2008.09.12 H.Nagamine ڼ�ߔԍ��w��ړ�
 * 2008.09.16 H.Nagamine ���ڔԍ��w��ړ�
 *  @@ T6 �ǉ�
 * Update 2013.10.21 ����n�Q�Ƌ@�\ (@20131021)
 */
public class CZControlTable extends JFrame {

    private boolean haita_flg = false;

    private final int T1 = 1;
    private final int T2 = 2;
    private final int T3 = 3;
    private final int T4 = 4;
    private final int T5 = 5;
    private final int T6 = 6;  //@@

    private CZNativeHikiage current_bt_set  = null;

    private Vector pv_data_shld = null;     //�V�����_�[�̃f�[�^
    private Vector pv_data_body = null;     //�{�f�B�[�̃f�[�^

    private JButton     grp_all_cp_button   = null;
    private JButton     grp_cp_button       = null;

    private JButton     recip_cp_button     = null;
    private JButton     recip_title_button  = null;

    private JButton     tbl_cp_button       = null;
    private JButton     tbl_title_button    = null;

    private JButton     t6LagCpButton_      = null;
    private JButton     t6LagSetButton_     = null;

    private JButton     t6MidCpButton_      = null;
    private JButton     t6MidSetButton_     = null;

    private JButton     t6ItemCpButton_     = null;
    private JButton     t6ItemSetButton_    = null;
    private JButton     t6ItemChgButton_    = null;

    private JButton     send_button         = null;
    private JButton     cancel_button       = null;
// add start 2008.09.12
    private RcpText     rcp_no_txt          = null;     // ���V�s
// add start 2008.09.12
// add start 2008.09.16
    private KoumokuText koumoku_no_txt      = null;     // ����
// add end 

    private Vector      table_title         = null;     //
    private CurrentTable    c_table         = null;     //�O���[�v
    private GroupeTable g_table             = null;     //���V�s
// add start 2008.09.12
    private JScrollPane rcp_pnl             = null;
// add end 2008.09.12
// add start 2008.09.16
    private JScrollPane kmk_pnl             = null;
// add end 2008.09.16

    private ValueTable  v_table             = null;     //����
    private T6LagTable  t6LagTable_         = null;     //T6�區��
    private T6MidTable  t6MidTable_         = null;     //T6������
    private T6ValTable  t6ValTable_         = null;     //T6����
    private Vector      t6Current_          = null;

    private CZControlTableSub setWin        = null;     //���ڐݒ�
    private T6SetWin setT6Win_              = null;     //T6���ڐݒ�

    private LimitWin    limitWin            = null;     //�����W
    private TitleWin    titleWin            = null;     //�^�C�g��

    private T6LagSetWin t6LagSetWin_        = null;
    private T6MidSetWin t6MidSetWin_        = null;
    private T6LimitWin  t6LimitWin_         = null;
    
    private CloseAlermWin closeAlermWin_    = null;

    private JLabel      ro_name_lab         = null;

    private CZControlTableCp cp_win         = null;     //�R�s�[����

//	private Timer       t                   = null;
//	private Timer       at                  = null;
//	private Timer       att                 = null;
//	private Timer       tcnt                = null;

	public Timer       t                   = null;
	public Timer       at                  = null;
	public Timer       att                 = null;
	public Timer       tcnt                = null;
	
	private int         tcount              = CZSystemDefine.ALERM_DIALOG_CLOSE_TIME/1000;
	
	/***** 2007.01.24 ADD *****/
	private boolean     Status_flg;
	
	private boolean     Button_flg;
	/**************************/

    //
    // �R���X�g���N�^
    //
    CZControlTable(){
        super();


        setTitle("����e�[�u���ݒ�");
// chg start 2008.09.10
//        setSize(960,740);
        setSize(960,760);
// chg end 2008.09.10
        setResizable(false);
        //setModal(true);

        addWindowListener(new WindowAdapter(){
            public void windowClosing(WindowEvent e){
				CZSystem.log("CZControlTable","����e�[�u���ݒ��ʃN���[�Y");
				setWin.setVisible(false);
                winClose(e);
            }
        });

        getContentPane().setLayout(null);
        // ����n�Q�Ƌ@�\    @20131021
        if(CZSystemDefine.REFERENCE_RUN == CZSystem.getRunLevel()){
            getContentPane().setBackground(CZSystemDefine.DEFAULT_REFERENCE_BACKGROUND_COL);
        }else{
            getContentPane().setBackground(CZSystemDefine.DEFAULT_BACKGROUND_COL);
        }
// add start 2008.09.12
        JLabel label = null;
// add end 2008.09.12
		String s = CZSystem.RoKetaChg(CZSystem.getRoName());	// 20050725 �F�F�\�������ύX
        ro_name_lab = new JLabel(s,JLabel.CENTER);
//        ro_name_lab = new JLabel(CZSystem.getRoName(),JLabel.CENTER);
        ro_name_lab.setBounds(20, 20, 100, 30);
        ro_name_lab.setLocale(new Locale("ja","JP"));
        ro_name_lab.setFont(new java.awt.Font("dialog", 0, 18));
        ro_name_lab.setBorder(new Flush3DBorder());
        ro_name_lab.setForeground(java.awt.Color.black);
        getContentPane().add(ro_name_lab);

        grp_all_cp_button = new JButton("�S�R�s�[");
// chg start 2008.09.10
//        grp_all_cp_button.setBounds(20, 170, 100, 24);
        grp_all_cp_button.setBounds(20, 190, 100, 24);
// chg end 2008.09.10
        grp_all_cp_button.setLocale(new Locale("ja","JP"));
        grp_all_cp_button.setFont(new java.awt.Font("dialog", 0, 18));
        grp_all_cp_button.setBorder(new Flush3DBorder());
        grp_all_cp_button.setForeground(java.awt.Color.black);
        grp_all_cp_button.addActionListener(new RoAllCopyButton());
        getContentPane().add(grp_all_cp_button);
        //�O���[�v�R�s�[
        grp_cp_button = new JButton("�R�s�[");
// chg start 2008.09.10
//        grp_cp_button.setBounds(140, 170, 100, 24);
        grp_cp_button.setBounds(140, 190, 100, 24);
// chg end 2008.09.10
        grp_cp_button.setLocale(new Locale("ja","JP"));
        grp_cp_button.setFont(new java.awt.Font("dialog", 0, 18));
        grp_cp_button.setBorder(new Flush3DBorder());
        grp_cp_button.setForeground(java.awt.Color.black);
        grp_cp_button.addActionListener(new GroupCopyButton());
        getContentPane().add(grp_cp_button);

// add start 2008.09.12
        label = new JLabel("���V�sNO  :",JLabel.CENTER);
        label.setBounds(260, 190, 100, 24);
        label.setLocale(new Locale("ja","JP"));
        label.setFont(new java.awt.Font("dialog", 0, 18));
        label.setForeground(java.awt.Color.black);
        getContentPane().add(label);

        rcp_no_txt = new RcpText();
        rcp_no_txt.setBounds(370, 190, 50, 24);
        rcp_no_txt.addActionListener(new RecipeAction());
        getContentPane().add(rcp_no_txt);

// add end 2008.09.12
        //���V�s
        //////////////////////////////////////////////////////////////////////
        recip_cp_button = new JButton("�R�s�[");
// chg start 2008.09.10
//        recip_cp_button.setBounds(20, 420, 100, 24);
        recip_cp_button.setBounds(20, 440, 100, 24);
// chg end 2008.09.10
        recip_cp_button.setLocale(new Locale("ja","JP"));
        recip_cp_button.setFont(new java.awt.Font("dialog", 0, 18));
        recip_cp_button.setBorder(new Flush3DBorder());
        recip_cp_button.setForeground(java.awt.Color.black);
        recip_cp_button.addActionListener(new RecipeCopyButton());
        getContentPane().add(recip_cp_button);

        recip_title_button = new JButton("�^�C�g��");
// chg start 2008.09.10
//        recip_title_button.setBounds(140, 420, 100, 24);
        recip_title_button.setBounds(140, 440, 100, 24);
// chg start 2008.09.10
        recip_title_button.setLocale(new Locale("ja","JP"));
        recip_title_button.setFont(new java.awt.Font("dialog", 0, 18));
        recip_title_button.setBorder(new Flush3DBorder());
        recip_title_button.setForeground(java.awt.Color.black);
        recip_title_button.addActionListener(new TitleButton());
        getContentPane().add(recip_title_button);
// add start 2008.09.16
        label = new JLabel("���ڔԍ�  :",JLabel.CENTER);
        label.setBounds(260, 440, 100, 24);
        label.setLocale(new Locale("ja","JP"));
        label.setFont(new java.awt.Font("dialog", 0, 18));
        label.setForeground(java.awt.Color.black);
        getContentPane().add(label);

        koumoku_no_txt = new KoumokuText();
        koumoku_no_txt.setBounds(370, 440, 50, 24);
        koumoku_no_txt.addActionListener(new KoumokuAction());
        getContentPane().add(koumoku_no_txt);

// add end 2008.09.16

        //�e�[�u��
        //////////////////////////////////////////////////////////////////////
        tbl_cp_button = new JButton("�R�s�[");
// chg start 2008.09.10
//        tbl_cp_button.setBounds(20, 670, 100, 24);
        tbl_cp_button.setBounds(20, 690, 100, 24);
// chg end 2008.09.10
        tbl_cp_button.setLocale(new Locale("ja","JP"));
        tbl_cp_button.setFont(new java.awt.Font("dialog", 0, 18));
        tbl_cp_button.setBorder(new Flush3DBorder());
        tbl_cp_button.setForeground(java.awt.Color.black);
        tbl_cp_button.addActionListener(new TableCopyButton());
        getContentPane().add(tbl_cp_button);

        tbl_title_button = new JButton("�����W");
// chg start 2008.09.10
//        tbl_title_button.setBounds(140, 670, 100, 24);
        tbl_title_button.setBounds(140, 690, 100, 24);
// chg end 2008.09.10
        tbl_title_button.setLocale(new Locale("ja","JP"));
        tbl_title_button.setFont(new java.awt.Font("dialog", 0, 18));
        tbl_title_button.setBorder(new Flush3DBorder());
        tbl_title_button.setForeground(java.awt.Color.black);
        tbl_title_button.addActionListener(new LimitButton());
        getContentPane().add(tbl_title_button);

        send_button = new JButton("�ݒ�ύX");
// chg start 2008.09.10
//        send_button.setBounds(260, 670, 100, 24);
        send_button.setBounds(260, 690, 100, 24);
// chg end 2008.09.10
        send_button.setLocale(new Locale("ja","JP"));
        send_button.setFont(new java.awt.Font("dialog", 0, 18));
        send_button.setBorder(new Flush3DBorder());
        send_button.setForeground(java.awt.Color.black);
        send_button.addActionListener(new SendButton());
        getContentPane().add(send_button);

        //////////////////////////////////////////////////////////////////////
        t6LagCpButton_ = new JButton("�R�s�[");
// chg start 2008.09.10
//        t6LagCpButton_.setBounds(580, 220, 100, 24);
        t6LagCpButton_.setBounds(580, 230, 100, 24);
// chg end 2008.09.10
        t6LagCpButton_.setLocale(new Locale("ja","JP"));
        t6LagCpButton_.setFont(new java.awt.Font("dialog", 0, 18));
        t6LagCpButton_.setBorder(new Flush3DBorder());
        t6LagCpButton_.setForeground(java.awt.Color.black);
        t6LagCpButton_.addActionListener(new T6LagCopyButton());
        getContentPane().add(t6LagCpButton_);
/*@@
        t6LagSetButton_ = new JButton("�區��");
        t6LagSetButton_.setBounds(700, 220, 100, 24);
        t6LagSetButton_.setLocale(new Locale("ja","JP"));
        t6LagSetButton_.setFont(new java.awt.Font("dialog", 0, 18));
        t6LagSetButton_.setBorder(new Flush3DBorder());
        t6LagSetButton_.setForeground(java.awt.Color.black);
        t6LagSetButton_.addActionListener(new T6LagSetButton());
        getContentPane().add(t6LagSetButton_);
@@*/
        t6MidCpButton_ = new JButton("�R�s�[");
// chg start 2008.09.10
//        t6MidCpButton_.setBounds(580, 420, 100, 24);
        t6MidCpButton_.setBounds(580, 440, 100, 24);
// chg end 2008.09.10
        t6MidCpButton_.setLocale(new Locale("ja","JP"));
        t6MidCpButton_.setFont(new java.awt.Font("dialog", 0, 18));
        t6MidCpButton_.setBorder(new Flush3DBorder());
        t6MidCpButton_.setForeground(java.awt.Color.black);
        t6MidCpButton_.addActionListener(new T6MidCopyButton());
        getContentPane().add(t6MidCpButton_);
/*@@
        t6MidSetButton_ = new JButton("������");
        t6MidSetButton_.setBounds(700, 420, 100, 24);
        t6MidSetButton_.setLocale(new Locale("ja","JP"));
        t6MidSetButton_.setFont(new java.awt.Font("dialog", 0, 18));
        t6MidSetButton_.setBorder(new Flush3DBorder());
        t6MidSetButton_.setForeground(java.awt.Color.black);
        t6MidSetButton_.addActionListener(new T6MidSetButton());
        getContentPane().add(t6MidSetButton_);
@@*/
        t6ItemCpButton_ = new JButton("�R�s�[");
// chg start 2008.09.10
//        t6ItemCpButton_.setBounds(580, 670, 100, 24);
        t6ItemCpButton_.setBounds(580, 690, 100, 24);
// chg end 2008.09.10
        t6ItemCpButton_.setLocale(new Locale("ja","JP"));
        t6ItemCpButton_.setFont(new java.awt.Font("dialog", 0, 18));
        t6ItemCpButton_.setBorder(new Flush3DBorder());
        t6ItemCpButton_.setForeground(java.awt.Color.black);
        t6ItemCpButton_.addActionListener(new T6ItemCopyButton());
        getContentPane().add(t6ItemCpButton_);

        t6ItemSetButton_ = new JButton("�����W");
// chg start 2008.09.10
//        t6ItemSetButton_.setBounds(700, 670, 100, 24);
        t6ItemSetButton_.setBounds(700, 690, 100, 24);
// chg end 2008.09.10
        t6ItemSetButton_.setLocale(new Locale("ja","JP"));
        t6ItemSetButton_.setFont(new java.awt.Font("dialog", 0, 18));
        t6ItemSetButton_.setBorder(new Flush3DBorder());
        t6ItemSetButton_.setForeground(java.awt.Color.black);
        t6ItemSetButton_.addActionListener(new T6LimitButton());
        getContentPane().add(t6ItemSetButton_);

        t6ItemChgButton_ = new JButton("�ݒ�ύX");
// chg start 2008.09.10
//        t6ItemChgButton_.setBounds(820, 670, 100, 24);
        t6ItemChgButton_.setBounds(820, 690, 100, 24);
// chg end 2008.09.10
        t6ItemChgButton_.setLocale(new Locale("ja","JP"));
        t6ItemChgButton_.setFont(new java.awt.Font("dialog", 0, 18));
        t6ItemChgButton_.setBorder(new Flush3DBorder());
        t6ItemChgButton_.setForeground(java.awt.Color.black);
        t6ItemChgButton_.addActionListener(new T6ItemChangeButton());
        getContentPane().add(t6ItemChgButton_);
/*@@
        cancel_button = new JButton("�I  ��");
        cancel_button.setBounds(820, 670, 100, 24);
//@@        cancel_button.setBounds(460, 670, 100, 24);
        cancel_button.setLocale(new Locale("ja","JP"));
        cancel_button.setFont(new java.awt.Font("dialog", 0, 18));
        cancel_button.setBorder(new Flush3DBorder());
        cancel_button.setForeground(java.awt.Color.black);
        cancel_button.addActionListener(new CancelButton());
        getContentPane().add(cancel_button);
@@*/
        //�O���[�v�e�[�u��
        c_table = new CurrentTable();
        JTableHeader tabHead = c_table.getTableHeader();
        tabHead.setReorderingAllowed(false);
        JScrollPane panel = new JScrollPane(c_table);
// chg start 2008.09.10
//        panel.setBounds(20, 60, 540, 100);
        panel.setBounds(20, 60, 540, 121);
// chg end 2008.09.10
        getContentPane().add(panel);

        //���V�s�e�[�u��
        g_table = new GroupeTable();
        tabHead = g_table.getTableHeader();
        tabHead.setReorderingAllowed(false);

// chg start 2008.09.12
//        panel = new JScrollPane(g_table);
        rcp_pnl = new JScrollPane(g_table);
// chg start 2008.09.10
//        panel.setBounds(20, 210, 540, 200);
//        panel.setBounds(20, 230, 540, 200);
        rcp_pnl.setBounds(20, 230, 540, 200);
// chg end 2008.09.10
//        getContentPane().add(panel);
        getContentPane().add(rcp_pnl);
// chg end 2008.09.12

        //���ڃe�[�u��
        v_table = new ValueTable();
        tabHead = v_table.getTableHeader();
        tabHead.setReorderingAllowed(false);
// chg start 2008.09.16
//        panel = new JScrollPane(v_table);
        kmk_pnl = new JScrollPane(v_table);
// chg start 2008.09.10
//        panel.setBounds(20, 460, 540, 200);
//        panel.setBounds(20, 480, 540, 200);
        kmk_pnl.setBounds(20, 480, 540, 200);
// chg end 2008.09.10
//        getContentPane().add(panel);
        getContentPane().add(kmk_pnl);
// chg end 2008.09.16

        //T6�區�ڃe�[�u�� @@
        t6LagTable_ = new T6LagTable();
        tabHead = t6LagTable_.getTableHeader();
        tabHead.setReorderingAllowed(false);
        panel = new JScrollPane(t6LagTable_);
        panel.setBounds(580, 60, 358, 150);
        getContentPane().add(panel);

        //T6�����ڃe�[�u�� @@
        t6MidTable_ = new T6MidTable();
        tabHead = t6MidTable_.getTableHeader();
        tabHead.setReorderingAllowed(false);
        panel = new JScrollPane(t6MidTable_);
// chg start 2008.09.10
//        panel.setBounds(580, 260, 358, 150);
        panel.setBounds(580, 280, 358, 150);
// chg end 2008.09.10
        getContentPane().add(panel);

        //T6���ڃe�[�u�� @@
        t6ValTable_ = new T6ValTable();
        tabHead = t6ValTable_.getTableHeader();
        tabHead.setReorderingAllowed(false);
        panel = new JScrollPane(t6ValTable_);
// chg start 2008.09.10
//        panel.setBounds(580, 460, 358, 200);
        panel.setBounds(580, 480, 358, 200);
// chg end 2008.09.10
        getContentPane().add(panel);

        //���ڐݒ���
        setWin = new CZControlTableSub();
        setWin.setVisible(false);
        //�����W���
        limitWin = new LimitWin();
        limitWin.setVisible(false);
        //���V�s�^�C�g�����
        titleWin = new TitleWin();
        titleWin.setVisible(false);

        //T6�區�ډ��
        t6LagSetWin_ = new T6LagSetWin();
        t6LagSetWin_.setVisible(false);

        //T6�����ډ��
        t6MidSetWin_ = new T6MidSetWin();
        t6MidSetWin_.setVisible(false);

        //T6���ډ��
        t6LimitWin_ = new T6LimitWin();
        t6LimitWin_.setVisible(false);

        //���ڐݒ���
        setT6Win_ = new T6SetWin();
        setT6Win_.setVisible(false);

        //�R�s�[���
        cp_win = new CZControlTableCp();
        
        //��ʃN���[�Y�x�����
        closeAlermWin_ = new CloseAlermWin();
        closeAlermWin_.setVisible(false);
        
        if( 0 != CZSystemDefine.TIMER_FLG ){
	        t = new Timer( CZSystemDefine.CT_TABLE_CLOSE_TIME, new AlermWin() );
	        tcnt = new Timer( 1000, new CountDown() );
	        at = new Timer( CZSystemDefine.ALERM_DIALOG_CLOSE_TIME, new CancelClose() );
	        att = new Timer( 10, new HaitaKaihou() );
		}

    }  //CZControlTable


    /**
     *
     *       ��ʃN���[�Y�A���[���pWindow
     *
     */
    public class CloseAlermWin extends JDialog {
		
		public JLabel       cnt_lab         = null;
		private JLabel      lab             = null;
		private JButton     cancel_button   = null;
		
	    //
	    // �R���X�g���N�^
	    //
	    CloseAlermWin(){
	        super();

	        setTitle("��ʃN���[�Y�x��");
	        setSize(400,150);
	        setLocation(600,500);
	        setResizable(false);
	        setModal(true);
	        
	        addWindowListener(new WindowAdapter(){
	            public void windowClosing(WindowEvent e){
	                AlermWinClose(e);
	            }
	        });

	        getContentPane().setLayout(null);
            // ����n�Q�Ƌ@�\    @20131021
            if(CZSystemDefine.REFERENCE_RUN == CZSystem.getRunLevel()){
                getContentPane().setBackground(CZSystemDefine.DEFAULT_REFERENCE_BACKGROUND_COL);
            }else{
                getContentPane().setBackground(CZSystemDefine.DEFAULT_BACKGROUND_COL);
            }

	        cancel_button = new JButton("����");
	        cancel_button.setBounds(150, 60, 100, 24);
	        cancel_button.setLocale(new Locale("ja","JP"));
	        cancel_button.setFont(new java.awt.Font("dialog", 0, 18));
	        cancel_button.setBorder(new Flush3DBorder());
	        cancel_button.setForeground(java.awt.Color.black);
	        cancel_button.addActionListener(new AlermClose());
	        getContentPane().add(cancel_button);
	        
			cnt_lab = new JLabel("");
			cnt_lab.setBounds(70, 10, 30, 30);
			cnt_lab.setLocale(new Locale("ja","JP"));
			cnt_lab.setFont(new java.awt.Font("dialog", 0, 18));
//			cnt_lab.setBorder(new Flush3DBorder());
			cnt_lab.setForeground(java.awt.Color.black);
			getContentPane().add(cnt_lab);	        

			lab = new JLabel("�b��ɉ�ʂ���܂�");
			lab.setBounds(100, 10, 250, 30);
			lab.setLocale(new Locale("ja","JP"));
			lab.setFont(new java.awt.Font("dialog", 0, 18));
			lab.setForeground(java.awt.Color.black);
			getContentPane().add(lab);	        
	    }
	}

	/********************************
	*
	* �J�E���g�_�E��
	*
	********************************/
	class CountDown implements ActionListener{
		public void actionPerformed( ActionEvent a ){
			
			tcount = tcount - 1;
			
			Integer i = new Integer( tcount );
			
			CZSystem.log("CZControlTable","�A���[����� �N���[�Y�܂�" + i + "�b");
			
			closeAlermWin_.cnt_lab.setText( i.toString() );
			
		}
	}

	/********************************
	*
	* �A���[�����
	*
	********************************/
	class AlermWin implements ActionListener{
		public void actionPerformed( ActionEvent a ){
			
			t.stop();
			at.restart();
			
			tcount = CZSystemDefine.ALERM_DIALOG_CLOSE_TIME/1000;
			
			CZSystem.log("CZControlTable","�A���[����� OPEN");
			
			tcnt.restart();
			
			closeAlermWin_.cnt_lab.setText("");
			
			closeAlermWin_.setVisible(true);

		}
	}


	/********************************
	*
	* �A���[����ʃN���[�Y
	*
	********************************/
	class AlermClose implements ActionListener{
		public void actionPerformed( ActionEvent a ){
			at.stop();
			tcnt.stop();
			t.restart();
			CZSystem.log("CZControlTable","�A���[����ʃI�[�v�����X�^�[�g�i�A���[�������j");
			CZSystem.log("CZControlTable","�A���[����ʃN���[�Y");
			closeAlermWin_.setVisible(false);
		}
	}

	/********************************
	*
	* ��ʃN���[�Y
	*
	********************************/
	class CancelClose implements ActionListener{
		public void actionPerformed( ActionEvent a ){
			at.stop();
			tcnt.stop();
			t.stop();
			limitWin.setVisible(false);
			titleWin.setVisible(false);
			t6LagSetWin_.setVisible(false);
			t6MidSetWin_.setVisible(false);
			t6LimitWin_.setVisible(false);
			setT6Win_.setVisible(false);
			setWin.setVisible(false);

			cp_win.ro_all_win.setVisible(false);
			cp_win.group_win.setVisible(false);
			cp_win.recipe_win.setVisible(false);
			cp_win.table_win.setVisible(false);
			cp_win.t6LagWin_.setVisible(false);
			cp_win.t6MidWin_.setVisible(false);
			cp_win.t6ItemWin_.setVisible(false);

			closeAlermWin_.setVisible(false);
			setVisible(false);
			att.restart();
		}
	}

	/********************************
	*
	* �r���J��
	*
	********************************/
	class HaitaKaihou implements ActionListener{
		public void actionPerformed( ActionEvent a ){
			putHaita();
			att.stop();
		}
	}

    //
    // �A���[����ʃN���[�Y
    //
    private void AlermWinClose(WindowEvent e){
        CZSystem.log("CZControlTable","AlermWinClose() " + e);
			at.stop();
			tcnt.stop();
			t.restart();
			CZSystem.log("CZControlTable","�A���[����ʃI�[�v�����X�^�[�g�i�~�j");
			CZSystem.log("CZControlTable","�A���[����ʃN���[�Y");
    }

	public boolean timerStart(){
		at.stop();
		t.restart();
		CZSystem.log("CZControlTable","�A���[����ʃI�[�v�����X�^�[�g�i���j���[�j");
		CZSystem.log("CZControlTable","�f�t�H���g�ݒ�");
	
		return true;
	}
	
	
    /**
    * �r���擾�v��
    */
    private boolean getHaita(){

        // ���Ɏ���Ă�ꍇ
        if(haita_flg) return true;

        String ro = CZSystem.getRoName();

        CZEventCL event = null;

        CZSystemQueue   que = new CZSystemQueue(20);
        CZEventAdapter  adp = new CZEventAdapter(que);
        CZEventSender.addCZEventListener(adp);

        boolean ret = CZSystem.CZGetControlExclusion(ro);

        haita_flg = false;

        if(!ret){
            CZEventSender.removeCZEventListener(adp);
            return false;
        }

        while(true){
            try{
                CZSystem.log("CZControlTable","getHaita() 1");
                event = (CZEventCL)que.waitObject();

                CZSystem.log("CZControlTable","getHaita() 2");
                // �r�������ȊO
                if(event.getEvent() != CZEventCL.CT_GET_HAITA) continue;
                CZSystem.log("CZControlTable","getHaita() 3");

                CZResult ev = (CZResult)event.getObject();

                CZSystem.log("CZControlTable","getHaita() 4");
                // �Ⴄ�F�̏ꍇ
                if(!ro.equals(ev.getRoban())) continue;

                CZSystem.log("CZControlTable","getHaita() 5");

                // �r���擾���s
                if(0 != ev.getStatus()){
                    CZSystem.log("CZControlTable","getHaita() 6");
                    CZEventSender.removeCZEventListener(adp);

                    CZSystemSysMsg msg = new CZSystemSysMsg();
                    msg.no = -1;
                    msg.message = CZSystem.getDateTime() + " ����e�[�u���r���擾���s [" + ev.getStatus() + "]";
                    CZSystem.sysMessage(msg);

                    return false;
                }

                CZSystem.log("CZControlTable","getHaita() 7");
                CZEventSender.removeCZEventListener(adp);
                haita_flg = true;
                {
                    CZSystemSysMsg msg = new CZSystemSysMsg();
                    msg.no = 0;
                    msg.message = CZSystem.getDateTime() + " ����e�[�u���r���擾���� [" + ev.getStatus() + "]";
                    CZSystem.sysMessage(msg);
                }
                return true;
            }
            catch(Exception e){
                CZSystem.handleException(e);
            }
        } //while end
    } //private boolean getHaita()

    //
    // �r���J���v��
    //
    private boolean putHaita(){

        String ro = CZSystem.getRoName();

        // ��ɊJ������l�ɕύX         01.03.27
        boolean ret = CZSystem.CZPutControlExclusion(ro);
        haita_flg = false;
        CZSystem.log("CZControlTable","putHaita() �r���J���v�� 2");

        return true;
    } //private boolean putHaita()


    //
    // ��ʃN���[�Y
    //
    private void winClose(WindowEvent e){
        CZSystem.log("CZControlTable","winClose() " + e);
	        if( 0 != CZSystemDefine.TIMER_FLG ){
		        t.stop();
		        at.stop();
		        att.stop();
		        tcnt.stop();
			}
        putHaita();
    }


    //
    // �f�t�H���g�ݒ�
    //
    public boolean setDefault(){
//@@        CZSystem.log("CZControlTable","setDefault()");

		String s = CZSystem.RoKetaChg(CZSystem.getRoName());	// 20050725 �F�F�\�������ύX
CZSystem.log("setDefault",CZSystem.getRoName());
CZSystem.log("setDefault",s);
        ro_name_lab.setText(s);
//        ro_name_lab.setText(CZSystem.getRoName());

        // @20131021 ����n�Q�Ƌ@�\
        if(CZSystemDefine.REFERENCE_RUN != CZSystem.getRunLevel()){  // �Q�Ƃ݂̂̏ꍇ�A�r�������͎��s���Ȃ�

        if(!getHaita()){
            Object msg[] = {"����e�[�u���r���擾",
                                "����ՁA���̒[����",
                                "�C�����ł�"};
            errorMsg(msg);
        }

        }  // @20131021

        pv_data_shld = null;                     //�V�����_�[�̃f�[�^
        pv_data_body = null;                     //�{�f�B�[�̃f�[�^

        current_bt_set = CZSystem.getBtSet();
        table_title    = CZSystem.getCtTitle();

        setCurrent();

        c_table.setRowSelectionInterval(0,0);
        g_table.setRowSelectionInterval(0,0);
        v_table.setRowSelectionInterval(0,0);

        /////////////////////////////////////////////////////////
        if(haita_flg){
            grp_all_cp_button.setEnabled(false);
            grp_cp_button.setEnabled(false);
            recip_cp_button.setEnabled(true);
            recip_title_button.setEnabled(true);
            tbl_cp_button.setEnabled(true);
            tbl_title_button.setEnabled(false);
        }
        else {
            grp_all_cp_button.setEnabled(false);
            grp_cp_button.setEnabled(false);
            recip_cp_button.setEnabled(false);
            recip_title_button.setEnabled(false);
            tbl_cp_button.setEnabled(false);
            tbl_title_button.setEnabled(false);
            t6LagCpButton_.setEnabled(false);
            t6MidCpButton_.setEnabled(false);
            t6ItemCpButton_.setEnabled(false);
            t6ItemSetButton_.setEnabled(false);
            t6ItemChgButton_.setEnabled(false);
        }

        if(CZSystemDefine.ADMIN_RUN == CZSystem.getRunLevel()){
            grp_all_cp_button.setEnabled(true);
            grp_cp_button.setEnabled(true);
        }
            recip_cp_button.setEnabled(true);
            recip_title_button.setEnabled(true);
            tbl_cp_button.setEnabled(true);
            tbl_title_button.setEnabled(true);
            t6LagCpButton_.setEnabled(true);
            t6MidCpButton_.setEnabled(true);
            t6ItemCpButton_.setEnabled(true);
            t6ItemSetButton_.setEnabled(true);
            t6ItemChgButton_.setEnabled(true);

        // ����n�Q�Ƌ@�\    @20131021
        if(CZSystemDefine.REFERENCE_RUN == CZSystem.getRunLevel()){
            grp_all_cp_button.setEnabled(false);
            grp_cp_button.setEnabled(false);
            recip_cp_button.setEnabled(false);
            recip_title_button.setEnabled(false);
            tbl_cp_button.setEnabled(false);
            tbl_title_button.setEnabled(false);
            //send_button.setEnabled(false);
            t6LagCpButton_.setEnabled(false);
            t6MidCpButton_.setEnabled(false);
            t6ItemCpButton_.setEnabled(false);
            t6ItemSetButton_.setEnabled(false);
            t6ItemChgButton_.setEnabled(false);
        }
        // @20131021

        return true;
    } //public boolean setDefault()


    //
    //  �o�u�f�[�^���ڂ������ꍇ
    //
    public boolean setDefault(Vector shld,Vector body){
//@@        CZSystem.log("CZControlTable","setDefault(Vector ,Vector)");

        boolean ret = this.setDefault();

        pv_data_shld = shld;                     //�V�����_�[�̃f�[�^
        pv_data_body = body;                     //�{�f�B�[�̃f�[�^

        return ret;
    }

    //
    //����l��ݒ肷��B
    //
    private boolean setCurrent(){

        Integer val     = null;
        CZSystemCtTitle title   = null;

        // �n��:T1
        val = new Integer(current_bt_set.getNo_youkai());
        c_table.setValueAt(val,0,1);
        for(int i = 0 ; i < table_title.size() ; i++){
            title = (CZSystemCtTitle)table_title.elementAt(i);
            if(T1 == title.g_no && current_bt_set.getNo_youkai() == title.r_no){
                c_table.setValueAt(title.title,0,2);
                break;
            }
        }

        // ����:T2
        val = new Integer(current_bt_set.getNo_hikiage());
        c_table.setValueAt(val,1,1);
        for(int i = 0 ; i < table_title.size() ; i++){
            title = (CZSystemCtTitle)table_title.elementAt(i);
            if(T2 == title.g_no && current_bt_set.getNo_hikiage() == title.r_no){
                c_table.setValueAt(title.title,1,2);
                break;
            }
        }

        // ��]:T3
        val = new Integer(current_bt_set.getNo_kaiten());
        c_table.setValueAt(val,2,1);
        for(int i = 0 ; i < table_title.size() ; i++){
            title = (CZSystemCtTitle)table_title.elementAt(i);
            if(T3 == title.g_no &&
                current_bt_set.getNo_kaiten() == title.r_no){
                c_table.setValueAt(title.title,2,2);
                break;
            }
        }

        // ��o:T4
        val = new Integer(current_bt_set.getNo_toridasi());
        c_table.setValueAt(val,3,1);
        for(int i = 0 ; i < table_title.size() ; i++){
            title = (CZSystemCtTitle)table_title.elementAt(i);
            if(T4 == title.g_no &&
                current_bt_set.getNo_toridasi() == title.r_no){
                c_table.setValueAt(title.title,3,2);
                break;
            }
        }

        // ����:T5
        val = new Integer(current_bt_set.getNo_aturyoku());
        c_table.setValueAt(val,4,1);
        for(int i = 0 ; i < table_title.size() ; i++){
            title = (CZSystemCtTitle)table_title.elementAt(i);
            if(T5 == title.g_no &&
                current_bt_set.getNo_aturyoku() == title.r_no){
                c_table.setValueAt(title.title,4,2);
                break;
            }
        }

        // �萔:T6
        val = new Integer(current_bt_set.getNo_teisu());
        c_table.setValueAt(val,5,1);
        for(int i = 0 ; i < table_title.size() ; i++){
            title = (CZSystemCtTitle)table_title.elementAt(i);
            if(T6 == title.g_no &&
                current_bt_set.getNo_teisu() == title.r_no){
                c_table.setValueAt(title.title,5,2);
                break;
            }
        }

        return true;
    } //private boolean setCurrent()


    //
    // ���b�Z�[�W�̕\��
    //
    private boolean errorMsg(Object msg[]){
        JOptionPane.showMessageDialog(null,msg,
                                    "����e�[�u��",
        JOptionPane.ERROR_MESSAGE);
        return true;
    }


    /**
     *
     *     �ݒ�ύX
     *
     */
    class SendButton implements ActionListener {
        public void actionPerformed(ActionEvent ev){

        if( 0 != CZSystemDefine.TIMER_FLG ){
			at.stop();
			t.restart();
			CZSystem.log("CZControlTable","�A���[����ʃI�[�v�����X�^�[�g�i�ݒ�ύX�I�[�v���j");
		}

            int g = c_table.getSelectedRow();
            int r = g_table.getSelectedRow();
            int n = v_table.getSelectedRow();

            if(0 > g) return;
            if(0 > r) return;
            if(0 > n) return;
            g++;
            r++;
            n++;

//@@            CZSystem.log("CZControlTable","SendButton [" + g + "][" + r + "][" + n + "]");

            Integer number = (Integer)v_table.getValueAt(v_table.getSelectedRow(),0);

            CZSystemCtName name =  null;
            name = CZSystem.getCtTbName(g,number.intValue());

            if(null == name) return;

//@@            CZSystem.log("CZControlTable","SendButton [" +
//@@                                    name.t_name + "][" + name.l_name + "][" + name.r_name+ "]");

            Integer group  = (Integer)c_table.getValueAt(c_table.getSelectedRow(),1);
            Integer recip  = (Integer)g_table.getValueAt(g_table.getSelectedRow(),0);

            boolean current = false;
            if(group.intValue() == recip.intValue()) current = true;

            setWin.setDefault(g,recip.intValue(),number.intValue(),name,
                                current,haita_flg,pv_data_shld,pv_data_body);

            setWin.setVisible(true);
            
            return ;
        }
    }  //class SendButton implements ActionListener


    /**
     *
     *      �����W
     *
     */
    class LimitButton implements ActionListener {
        public void actionPerformed(ActionEvent ev){

        if( 0 != CZSystemDefine.TIMER_FLG ){
			at.stop();
			t.restart();
			CZSystem.log("CZControlTable","�A���[����ʃI�[�v�����X�^�[�g�i�����W�I�[�v���j");
		}
		
            int g = c_table.getSelectedRow();
            int r = g_table.getSelectedRow();
            int n = v_table.getSelectedRow();

            if(0 > g) return;
            if(0 > r) return;
            if(0 > n) return;
            g++;
            r++;
            n++;

//@@            CZSystem.log("CZControlTable","LimitButton [" + g + "][" + r + "][" + n + "]");

            Integer number = (Integer)v_table.getValueAt(v_table.getSelectedRow(),0);

            CZSystemCtName name =  null;
            name = CZSystem.getCtTbName(g,number.intValue());

            if(null == name) return;

//@@            CZSystem.log("CZControlTable","LimitButton [" +
//@@                            name.t_name + "][" + name.l_name + "][" + name.r_name+ "]");

            Integer group  = (Integer)c_table.getValueAt(c_table.getSelectedRow(),1);
            Integer recip  = (Integer)g_table.getValueAt(g_table.getSelectedRow(),0);

            boolean current = false;
            if(group.intValue() == recip.intValue()) current = true;

            limitWin.setDefault(name);

            limitWin.setVisible(true);

            return ;
        }
    } //class LimitButton implements ActionListener


    /**
     *
     *   �e�[�u���^�C�g���ύX
     *
     */
    class TitleButton implements ActionListener {
        public void actionPerformed(ActionEvent ev){

        if( 0 != CZSystemDefine.TIMER_FLG ){
			at.stop();
			t.restart();
			CZSystem.log("CZControlTable","�A���[����ʃI�[�v�����X�^�[�g�i�^�C�g���I�[�v���j");
		}
		
            int g = c_table.getSelectedRow();
            int r = g_table.getSelectedRow();

            if(0 > g) return;
            if(0 > r) return;
            g++;
            r++;

//@@            CZSystem.log("CZControlTable","TitleButton [" + g + "][" + r + "]");

            String group = (String)c_table.getValueAt(c_table.getSelectedRow(),0);
            String title = (String)g_table.getValueAt(g_table.getSelectedRow(),1);
            if(null == title){
                title = new String("");
            }

            titleWin.setDefault(g,group,r,title);

            titleWin.setVisible(true);

            setDefault();
            c_table.setRowSelectionInterval(0,0);
            c_table.setRowSelectionInterval(1,1);
            c_table.setRowSelectionInterval(2,2);
            c_table.setRowSelectionInterval(3,3);
            c_table.setRowSelectionInterval(4,4);
            c_table.setRowSelectionInterval(5,5);
            c_table.setRowSelectionInterval(g-1,g-1);
            g_table.setRowSelectionInterval(r-1,r-1);
            g_table.repaint();

            return ;
        }
    } //class TitleButton implements ActionListener


    /*
    *   �區�ڕύX
    */
    class T6LagSetButton implements ActionListener {
        public void actionPerformed(ActionEvent ev){

            int g = c_table.getSelectedRow();
            int r = g_table.getSelectedRow();
            int l = t6LagTable_.getSelectedRow();

            if(0 > g) return;
            if(0 > r) return;
            if(0 > l) return;
            g++;
            r++;
            l++;

//@@            CZSystem.log("CZControlTable","T6LagSetButton [" + g + "][" + r + "][" + l + "]");

            String group   = (String)c_table.getValueAt(c_table.getSelectedRow(),0);
            String title   = (String)g_table.getValueAt(g_table.getSelectedRow(),1);
            String lagName = (String)t6LagTable_.getValueAt(t6LagTable_.getSelectedRow(),1);
            if(null == title){
                title = new String("");
            }

            t6LagSetWin_.setDefault(g,group,r,title,l,lagName);
            t6LagSetWin_.setVisible(true);

            setDefault();
            c_table.setRowSelectionInterval(0,0);
            c_table.setRowSelectionInterval(1,1);
            c_table.setRowSelectionInterval(2,2);
            c_table.setRowSelectionInterval(3,3);
            c_table.setRowSelectionInterval(4,4);
            c_table.setRowSelectionInterval(5,5);
            c_table.setRowSelectionInterval(g-1,g-1);
            g_table.setRowSelectionInterval(r-1,r-1);
            t6LagTable_.setRowSelectionInterval(l-1,l-1);
            t6LagTable_.repaint();

            return ;
        }
    } //class T6LagSetButton implements ActionListener

    /*
    *   �����ڕύX
    */
    class T6MidSetButton implements ActionListener {
        public void actionPerformed(ActionEvent ev){

            int g = c_table.getSelectedRow();
            int r = g_table.getSelectedRow();
            int l = t6LagTable_.getSelectedRow();
            int m = t6MidTable_.getSelectedRow();

            if(0 > g) return;
            if(0 > r) return;
            if(0 > l) return;
            if(0 > m) return;
            g++;
            r++;
            l++;
            m++;

//@@            CZSystem.log("CZControlTable","TitleButton [" +
//@@                                g + "][" + r + "][" + l + "][" + m + "]");

            String group   = (String)c_table.getValueAt(c_table.getSelectedRow(),0);
            String title   = (String)g_table.getValueAt(g_table.getSelectedRow(),1);
            String lagName = (String)t6LagTable_.getValueAt(t6LagTable_.getSelectedRow(),1);
            String midName = (String)t6MidTable_.getValueAt(t6MidTable_.getSelectedRow(),1);
            if(null == title){
                title = new String("");
            }

            t6MidSetWin_.setDefault(g, group, r, title, l, lagName, m, midName);

            t6MidSetWin_.setVisible(true);

            setDefault();
            c_table.setRowSelectionInterval(0,0);
            c_table.setRowSelectionInterval(1,1);
            c_table.setRowSelectionInterval(2,2);
            c_table.setRowSelectionInterval(3,3);
            c_table.setRowSelectionInterval(4,4);
            c_table.setRowSelectionInterval(5,5);
            c_table.setRowSelectionInterval(g-1,g-1);
            g_table.setRowSelectionInterval(r-1,r-1);
            t6LagTable_.setRowSelectionInterval(l-1,l-1);
            t6MidTable_.setRowSelectionInterval(m-1,m-1);
            t6MidTable_.repaint();

            return ;
        }
    } //class T6MidSetButton implements ActionListener

    /*
    *
    * T6�����W
    *
    */
    class T6LimitButton implements ActionListener {
        public void actionPerformed(ActionEvent ev){

            int gNo = c_table.getSelectedRow();
            int rNo = g_table.getSelectedRow();
            int lNo = t6LagTable_.getSelectedRow();
            int mNo = t6MidTable_.getSelectedRow();
            int iNo = t6ValTable_.getSelectedRow();

            if(0 > gNo) return;
            if(0 > rNo) return;
            if(0 > lNo) return;
            if(0 > mNo) return;
            if(0 > iNo) return;
            gNo++;
            rNo++;
            lNo++;
            mNo++;
            iNo++;

//@@            CZSystem.log("CZControlTable","LimitButton [" +
//@@                     gNo + "][" + rNo + "][" + lNo + "][" + mNo + "][" + iNo + "]");

            Integer number = (Integer)t6ValTable_.getValueAt(t6ValTable_.getSelectedRow(),0);

            CZSystemCtT6Name name =  null;
            name = CZSystem.getCtT6Name(gNo, lNo, mNo, number.intValue());

            Integer group  = (Integer)c_table.getValueAt(c_table.getSelectedRow(),1);
            Integer recip  = (Integer)g_table.getValueAt(g_table.getSelectedRow(),0);

            boolean current = false;
            if(group.intValue() == recip.intValue()) current = true;

            if (t6LimitWin_.setDefault(name)) {
                t6LimitWin_.setVisible(true);
//@@                CZSystem.log("CZControlTable","T6LimitWin show.");
            } else {
                CZSystem.log("CZControlTable","T6LimitWin Data nothing.");
            }
            return ;
        }
    } //class T6LimitButton implements ActionListener

    /**
     *
     *      �I��
     *
     */
    class CancelButton implements ActionListener {
        public void actionPerformed(ActionEvent ev){
            setVisible(false);
            putHaita();
        }
    } //class CancelButton implements ActionListener

    /**
     *
     *   �S�R�s�[
     *
     */
    class RoAllCopyButton implements ActionListener {
        public void actionPerformed(ActionEvent ev){

        if( 0 != CZSystemDefine.TIMER_FLG ){
			at.stop();
			t.restart();
			CZSystem.log("CZControlTable","�A���[����ʃI�[�v�����X�^�[�g�i�S�R�s�[�j");
		}
		
            cp_win.roAllCopy();

            setDefault();
            c_table.setRowSelectionInterval(0,0);
            c_table.setRowSelectionInterval(1,1);
            c_table.setRowSelectionInterval(2,2);
            c_table.setRowSelectionInterval(3,3);
            c_table.setRowSelectionInterval(4,4);
            c_table.setRowSelectionInterval(5,5);
            return ;
        }
    } //class RoAllCopyButton implements ActionListener


    /**
     *
     *   �O���[�v�R�s�[
     *
     */
    class GroupCopyButton implements ActionListener {
        public void actionPerformed(ActionEvent ev){

        if( 0 != CZSystemDefine.TIMER_FLG ){
			at.stop();
			t.restart();
			CZSystem.log("CZControlTable","�A���[����ʃI�[�v�����X�^�[�g�i�O���[�v�R�s�[�j");
		}
		
            int g = c_table.getSelectedRow();

            if(0 > g) return;
            g++;

//@@            CZSystem.log("CZControlTable","GroupCopyButton [" + g + "]");

            String group = (String)c_table.getValueAt(c_table.getSelectedRow(),0);

            cp_win.groupCopy(g,group);

            setDefault();
            c_table.setRowSelectionInterval(0,0);
            c_table.setRowSelectionInterval(1,1);
            c_table.setRowSelectionInterval(2,2);
            c_table.setRowSelectionInterval(3,3);
            c_table.setRowSelectionInterval(5,5);
            c_table.setRowSelectionInterval(g-1,g-1);

            return ;
        }
    } //class GroupCopyButton implements ActionListener

    /**
     *
     *   ���V�s�[�R�s�[
     *
     */
    class RecipeCopyButton implements ActionListener {
        public void actionPerformed(ActionEvent ev){

        if( 0 != CZSystemDefine.TIMER_FLG ){
			at.stop();
			t.restart();
			CZSystem.log("CZControlTable","�A���[����ʃI�[�v�����X�^�[�g�i���V�s�[�R�s�[�j");
		}
		
            int g = c_table.getSelectedRow();
            int r = g_table.getSelectedRow();

            if(0 > g) return;
            if(0 > r) return;
            g++;
            r++;

//@@            CZSystem.log("CZControlTable","RecipeCopyButton [" + g + "][" + r + "]");

            String group = (String)c_table.getValueAt(c_table.getSelectedRow(),0);
            String title = (String)g_table.getValueAt(g_table.getSelectedRow(),1);
            if(null == title){
                title = new String("");
            }

            cp_win.recipeCopy(g,group,r,title);

            setDefault();
            c_table.setRowSelectionInterval(0,0);
            c_table.setRowSelectionInterval(1,1);
            c_table.setRowSelectionInterval(2,2);
            c_table.setRowSelectionInterval(3,3);
            c_table.setRowSelectionInterval(4,4);
            c_table.setRowSelectionInterval(5,5);
            c_table.setRowSelectionInterval(g-1,g-1);
            g_table.setRowSelectionInterval(r-1,r-1);
            g_table.repaint();

            return ;
        }
    } //class RecipeCopyButton implements ActionListener

    /**
     *
     *   �e�[�u���R�s�[
     *
     */
    class TableCopyButton implements ActionListener {
        public void actionPerformed(ActionEvent ev){

        if( 0 != CZSystemDefine.TIMER_FLG ){
			at.stop();
			t.restart();
			CZSystem.log("CZControlTable","�A���[����ʃI�[�v�����X�^�[�g�i�e�[�u���R�s�[�j");
		}
		
             int g = c_table.getSelectedRow();
            int r = g_table.getSelectedRow();
//2011.04.12 Y.K start
            int select_row[] = v_table.getSelectedRows();
			String item;
			int iCnt = 0;
            for(int i = 0 ; i < select_row.length ; i++) {
	            item  = (String)v_table.getValueAt(select_row[i],1);
	            if(null != item){
					if (item != "")
					{
		                iCnt++;
					}
	            }
			}

//            CZSystem.log("CZControlTable TableCopyButton","iCnt [" + iCnt + "]");

            int row_list[] = new int[iCnt];
			String item_List[] = new String[iCnt];
			int		iSetCnt = 0;
            for(int i = 0 ; i < select_row.length ; i++) {
//                CZSystem.log("CZControlTable TableCopyButton","actionPerformed [" + i +
//                    "][" + select_row[i] + "]");
				
				item  = (String)v_table.getValueAt(select_row[i],1);
				if (item != null)
				{
					if (item != "")
					{
						row_list[iSetCnt] = select_row[i] + 1;
			            item_List[iSetCnt]  = (String)v_table.getValueAt(select_row[i],1);
			            if(null == item_List[iSetCnt]){
			                item_List[iSetCnt] = "";
			            }


//		                CZSystem.log("CZControlTable TableCopyButton 2","actionPerformed [" + iSetCnt +
//		                    "][" + row_list[iSetCnt] + "][" + item_List[iSetCnt] + "]");
						iSetCnt++;
					}
				}
			}
//2011.04.12 Y.K end

            if(0 > g) return;
            if(0 > r) return;
            if(1 > iSetCnt) return;  //if (0 > v) 2011.04.12 Y.K

            g++;
            r++;
//2011.04.12 Y.K            v++;


//@@            CZSystem.log("CZControlTable","TitleButton [" + g + "][" + r + "]");

            String group = (String)c_table.getValueAt(c_table.getSelectedRow(),0);
            String title = (String)g_table.getValueAt(g_table.getSelectedRow(),1);
//2011.04.12 Y.K cut            String item  = (String)v_table.getValueAt(v_table.getSelectedRow(),1);

            if(null == title){
//2011.04.12 Y.K Start
//                title = new String("");
                Object msg[] = {"����e�[�u��",
                                "�e�[�u�������݂��܂���I�I",
                                ""};
                errorMsg(msg);
				return;
//2011.04.12 Y.K End
            }

			CZSystem.log("CZControlTable","title[" + title +  "][" + title.trim() + "][" + title.trim().length() + "]");

            if(0 == title.trim().length()){
//2011.04.12 Y.K Start
//                title = new String("");
                Object msg[] = {"����e�[�u��",
                                "�e�[�u�������݂��܂���I�I",
                                ""};
                errorMsg(msg);
				return;
//2011.04.12 Y.K End
            }

//2011.04.12 Y.K Start
//            if(null == item){
//                item = new String("");
//            }
//            cp_win.tableCopy(g,group,r,title,v,item);
            cp_win.tableCopy(g,group,r,title,row_list,item_List);
////2011.04.12 Y.K end

            setDefault();
            c_table.setRowSelectionInterval(0,0);
            c_table.setRowSelectionInterval(1,1);
            c_table.setRowSelectionInterval(2,2);
            c_table.setRowSelectionInterval(3,3);
            c_table.setRowSelectionInterval(4,4);
            c_table.setRowSelectionInterval(5,5);
            c_table.setRowSelectionInterval(g-1,g-1);
            g_table.setRowSelectionInterval(r-1,r-1);
            g_table.repaint();

            return ;
        }
    } //class TableCopyButton implements ActionListener


    /*
    *   �區�ڃR�s�[
    */
    class T6LagCopyButton implements ActionListener {
        public void actionPerformed(ActionEvent ev){

        if( 0 != CZSystemDefine.TIMER_FLG ){
			at.stop();
			t.restart();
			CZSystem.log("CZControlTable","�A���[����ʃI�[�v�����X�^�[�g�i�區�ڃR�s�[�j");
		}

            int g = c_table.getSelectedRow();
            int r = g_table.getSelectedRow();
            int l = t6LagTable_.getSelectedRow();

            if(0 > g) return;
            if(0 > r) return;
            if(0 > l) return;
            g++;
            r++;
            l++;

//@@            CZSystem.log("CZControlTable","T6LagCopyButton [" +
//@@                                            g + "][" + r + "][" + l + "]");

            String group   = (String)c_table.getValueAt(c_table.getSelectedRow(),0);
            String recip   = (String)g_table.getValueAt(g_table.getSelectedRow(),1);
            String lagName = (String)t6LagTable_.getValueAt(t6LagTable_.getSelectedRow(),1);

            if(null == recip){
                recip = new String("");
            }

            if(null == lagName){
                lagName = new String("");
            }

            cp_win.t6LagCopy( g, group, r, recip, l, lagName );

            setDefault();
            c_table.setRowSelectionInterval(0,0);
            c_table.setRowSelectionInterval(1,1);
            c_table.setRowSelectionInterval(2,2);
            c_table.setRowSelectionInterval(3,3);
            c_table.setRowSelectionInterval(4,4);
            c_table.setRowSelectionInterval(5,5);
            c_table.setRowSelectionInterval(g-1,g-1);
            g_table.setRowSelectionInterval(r-1,r-1);
            t6LagTable_.setRowSelectionInterval(l-1,l-1);
            t6LagTable_.repaint();

            return ;
        }
    } //class T6LagCopyButton implements ActionListener

    /*
    *   �����ڃR�s�[
    */
    class T6MidCopyButton implements ActionListener {
        public void actionPerformed(ActionEvent ev){

        if( 0 != CZSystemDefine.TIMER_FLG ){
			at.stop();
			t.restart();
			CZSystem.log("CZControlTable","�A���[����ʃI�[�v�����X�^�[�g�i�����ڃR�s�[�j");
		}

            int g = c_table.getSelectedRow();
            int r = g_table.getSelectedRow();
            int l = t6LagTable_.getSelectedRow();
            int m = t6MidTable_.getSelectedRow();

            if(0 > g) return;
            if(0 > r) return;
            if(0 > l) return;
            if(0 > m) return;
            g++;
            r++;
            l++;
            m++;

//@@            CZSystem.log("CZControlTable","T6MidCopyButton [" +
//@@                                g + "][" + r + "][" + l + "][" + m + "]");

            String group   = (String)c_table.getValueAt(c_table.getSelectedRow(),0);
            String recip   = (String)g_table.getValueAt(g_table.getSelectedRow(),1);
            String lagName = (String)t6LagTable_.getValueAt(t6LagTable_.getSelectedRow(),1);
            String midName = (String)t6MidTable_.getValueAt(t6MidTable_.getSelectedRow(),1);

            if(null == recip){
                recip = new String("");
            }

            if(null == lagName){
                lagName = new String("");
            }

            if(null == midName){
                midName = new String("");
            }

            cp_win.t6MidCopy( g, group, r, recip, l, lagName, m, midName );

            setDefault();
            c_table.setRowSelectionInterval(0,0);
            c_table.setRowSelectionInterval(1,1);
            c_table.setRowSelectionInterval(2,2);
            c_table.setRowSelectionInterval(3,3);
            c_table.setRowSelectionInterval(4,4);
            c_table.setRowSelectionInterval(5,5);
            c_table.setRowSelectionInterval(g-1,g-1);
            g_table.setRowSelectionInterval(r-1,r-1);
            t6LagTable_.setRowSelectionInterval(l-1,l-1);
            t6MidTable_.setRowSelectionInterval(m-1,m-1);
            t6MidTable_.repaint();

            return ;
        }
    } //class T6MidCopyButton implements ActionListener

    /*
    *   ���ڃR�s�[
    */
    class T6ItemCopyButton implements ActionListener {
        public void actionPerformed(ActionEvent ev){

        if( 0 != CZSystemDefine.TIMER_FLG ){
			at.stop();
			t.restart();
			CZSystem.log("CZControlTable","�A���[����ʃI�[�v�����X�^�[�g�i���ڃR�s�[�j");
		}

            int gNo = c_table.getSelectedRow();
            int rNo = g_table.getSelectedRow();
            int lNo = t6LagTable_.getSelectedRow();
            int mNo = t6MidTable_.getSelectedRow();
            int iNo = t6ValTable_.getSelectedRow();

            if(0 > gNo) return;
            if(0 > rNo) return;
            if(0 > lNo) return;
            if(0 > mNo) return;
            if(0 > iNo) return;
            gNo++;
            rNo++;
            lNo++;
            mNo++;
            iNo++;

//@@            CZSystem.log("CZControlTable","T6ItemCopyButton [" +
//@@                                gNo + "][" + rNo + "][" + lNo + "][" + mNo + "][" + iNo + "]");

            String group   = (String)c_table.getValueAt(c_table.getSelectedRow(),0);
            String recip   = (String)g_table.getValueAt(g_table.getSelectedRow(),1);
            String lagName = (String)t6LagTable_.getValueAt(t6LagTable_.getSelectedRow(),1);
            String midName = (String)t6MidTable_.getValueAt(t6MidTable_.getSelectedRow(),1);
            String itmName = (String)t6ValTable_.getValueAt(t6ValTable_.getSelectedRow(),1);

            if(null == recip){
                recip = new String("");
            }

            if(null == lagName){
                lagName = new String("");
            }

            if(null == midName){
                midName = new String("");
            }

            if(null == itmName){
                itmName = new String("");
            }

            cp_win.t6ItemCopy(gNo, group, rNo, recip,
                lNo, lagName, mNo, midName, iNo, itmName);

            setDefault();
            c_table.setRowSelectionInterval(0,0);
            c_table.setRowSelectionInterval(1,1);
            c_table.setRowSelectionInterval(2,2);
            c_table.setRowSelectionInterval(3,3);
            c_table.setRowSelectionInterval(4,4);
            c_table.setRowSelectionInterval(5,5);
            c_table.setRowSelectionInterval(gNo-1,gNo-1);
            g_table.setRowSelectionInterval(rNo-1,rNo-1);
            t6LagTable_.setRowSelectionInterval(lNo-1,lNo-1);
            t6MidTable_.setRowSelectionInterval(mNo-1,mNo-1);
            t6ValTable_.setRowSelectionInterval(iNo-1,iNo-1);
            t6ValTable_.repaint();

            return ;
        }
    }  //class T6ItemCopyButton implements ActionListener

    /**
     *     T6�ݒ�ύX       //@@
     */
    class T6ItemChangeButton implements ActionListener {
        public void actionPerformed(ActionEvent ev){

        if( 0 != CZSystemDefine.TIMER_FLG ){
			at.stop();
			t.restart();
			CZSystem.log("CZControlTable","�A���[����ʃI�[�v�����X�^�[�g�iT6�ݒ�ύX�j");
		}

            int gNo = c_table.getSelectedRow();
            int rNo = g_table.getSelectedRow();
            int lNo = t6LagTable_.getSelectedRow();
            int mNo = t6MidTable_.getSelectedRow();

            if(0 > gNo) return;
            if(0 > rNo) return;
            if(0 > lNo) return;
            if(0 > mNo) return;
            gNo++;
            rNo++;
            lNo++;
            mNo++;

//@@            CZSystem.log("CZControlTable","T6ItemChangeButton [" +
//@@                gNo + "][" + rNo + "]["+ "][" + lNo + "]["+ "][" + mNo + "]");

            Integer group  = (Integer)c_table.getValueAt(c_table.getSelectedRow(),1);
            Integer recip  = (Integer)g_table.getValueAt(g_table.getSelectedRow(),0);

            boolean current = false;
            /***** 2007.01.24 ADD *****/
//            if(group.intValue() == recip.intValue()) current = true;
            if(group.intValue() == recip.intValue()){
				current = true;
                t6Current_ = CZSystem.getCtT6Tb(6, rNo, lNo, mNo, current);  /* �J�����g�\�� */
                Button_flg = true;
            } else {
                t6Current_ = CZSystem.getCtT6Tb(6, rNo, lNo, mNo, current);  /* �}�X�^�[�\�� */
                Button_flg = false;
            }

            CZSystem.log("CZControlTable","Current or Master [" + current + "]");

            Status_flg = current;
            /**************************/

            setT6Win_.setDefault( gNo, rNo, lNo, mNo );

            setT6Win_.setVisible(true);
            return ;
        }
    } //class T6ItemChangeButton implements ActionListener

    /**
     *
     *   �J�����g�̐���O���[�v�ꗗ
     *
     */
    class CurrentTable extends JTable {

        private CtTblMdl model = null;

        CurrentTable(){
            super();


            try{
                setName("CurrentTable");
                setBounds(0, 0, 200, 200);
                setAutoCreateColumnsFromModel(true);
                setSelectionMode(ListSelectionModel.SINGLE_SELECTION);

                setLocale(new Locale("ja","JP"));
                setFont(new java.awt.Font("dialog", 0, 12));
                setRowHeight(17);

                model = new CtTblMdl();
                setModel(model);

                DefaultTableColumnModel cmdl = (DefaultTableColumnModel)getColumnModel();
                TableColumn  colum = null;

                // �O���[�v
                colum = cmdl.getColumn(0);
                colum.setMaxWidth(60);
                colum.setMinWidth(60);
                colum.setWidth(60);

                // ���V�s�[No
                colum = cmdl.getColumn(1);
                colum.setMaxWidth(40);
                colum.setMinWidth(40);
                colum.setWidth(40);

                // �^�C�g��
                colum = cmdl.getColumn(2);
            }
            catch (Throwable e) {
                CZSystem.handleException(e);
            }
        }

        //
        //      �s��I��������
        //      ���V�s�A���ڃe�[�u�����������
        public void valueChanged(ListSelectionEvent e){
            super.valueChanged(e);

//@@            CZSystem.log("CZControlTable","CurrentTable valueChanged [" +
//@@                 getSelectedRow() + "][" + getSelectedColumn() + "]");

            if(0 > getSelectedRow()) return;

            Integer rec = (Integer)getValueAt(getSelectedRow(),1);

            g_table.setData(getSelectedRow()+1,rec.intValue());     //�O���[�v
            // T6��I�������ꍇ�́A�區�ځA�����ځA���ڂ���ւ���B
            // T1����T5�p�̃{�^���́A�g�p�s�ɂ���B
            if (5 != getSelectedRow()) {
                v_table.setData(getSelectedRow()+1);                    //����
                v_table.setRowSelectionInterval(0,0);
                v_table.setVisible(true);
                t6LagTable_.setVisible(false);
                t6MidTable_.setVisible(false);
                t6ValTable_.setVisible(false);

                tbl_cp_button.setEnabled(true);
                tbl_title_button.setEnabled(true);
                send_button.setEnabled(true);
                t6LagCpButton_.setEnabled(false);
                t6MidCpButton_.setEnabled(false);
                t6ItemCpButton_.setEnabled(false);
                t6ItemSetButton_.setEnabled(false);
                t6ItemChgButton_.setEnabled(false);
//@@                t6LagSetButton_.setEnabled(false);
//@@                t6MidSetButton_.setEnabled(false);

                // ����n�Q�Ƌ@�\    @20131021
                if(CZSystemDefine.REFERENCE_RUN == CZSystem.getRunLevel()){
                    tbl_cp_button.setEnabled(false);
                    tbl_title_button.setEnabled(false);
                    // send_button.setEnabled(false);
                    t6LagCpButton_.setEnabled(false);
                    t6MidCpButton_.setEnabled(false);
                    t6ItemCpButton_.setEnabled(false);
                    t6ItemSetButton_.setEnabled(false);
                    // t6ItemChgButton_.setEnabled(false);
                }
                // @20131021
            } else {
                //�抸�������V�s�ɑΉ�����區�ځA�����ڂ����o���B
                v_table.setVisible(false);
                t6LagTable_.setVisible(true);
                t6MidTable_.setVisible(true);
                t6ValTable_.setVisible(true);
                t6LagTable_.setData(getSelectedRow()+1,rec.intValue());
                Integer il= (Integer)t6LagTable_.getValueAt(0,0);
                t6MidTable_.setData(getSelectedRow()+1,rec.intValue(),il.intValue());
                Integer im= (Integer)t6MidTable_.getValueAt(0,0);
                t6ValTable_.setData(getSelectedRow()+1,il.intValue(),im.intValue());

                tbl_cp_button.setEnabled(false);
                tbl_title_button.setEnabled(false);
                send_button.setEnabled(false);
                t6LagCpButton_.setEnabled(true);
                t6MidCpButton_.setEnabled(true);
                t6ItemCpButton_.setEnabled(true);
                t6ItemSetButton_.setEnabled(true);
                t6ItemChgButton_.setEnabled(true);
//@@                t6LagSetButton_.setEnabled(true);
//@@                t6MidSetButton_.setEnabled(true);

                // ����n�Q�Ƌ@�\    @20131021
                if(CZSystemDefine.REFERENCE_RUN == CZSystem.getRunLevel()){
                    // send_button.setEnabled(false);
                    t6LagCpButton_.setEnabled(false);
                    t6MidCpButton_.setEnabled(false);
                    t6ItemCpButton_.setEnabled(false);
                    t6ItemSetButton_.setEnabled(false);
                    // t6ItemChgButton_.setEnabled(false);
                }
                // @20131021
            }
        }
    }  //class CurrentTable extends JTable

    /**
     *
     *       ����O���[�v�e�[�u���N���X�F���f��
     *
     */
    public class CtTblMdl extends AbstractTableModel {

        final   int TBL_COL = 3;
        private int TBL_ROW = 6;    //@@ 5 -> 6

        final String[] names = {"�O���[�v", " # " ,"�^�C�g��"};

        private Object  data[][];

        CtTblMdl(){
            super();

            data = new Object[TBL_ROW][TBL_COL];

            data[0][0] = new String("�n��:T1");
            data[1][0] = new String("����:T2");
            data[2][0] = new String("��]:T3");
            data[3][0] = new String("��o:T4");
            data[4][0] = new String("����:T5");
            data[5][0] = new String("�萔:T6");     //@@

            data[0][1] = new Integer(0);
            data[1][1] = new Integer(0);
            data[2][1] = new Integer(0);
            data[3][1] = new Integer(0);
            data[4][1] = new Integer(0);
            data[5][1] = new Integer(0);            //@@

            data[0][2] = new String("0#########1#########2#########3#########4#########5#########6###");
            data[1][2] = new String("0#########1#########2#########3#########4#########5#########6###");
            data[2][2] = new String("0#########1#########2#########3#########4#########5#########6###");
            data[3][2] = new String("0#########1#########2#########3#########4#########5#########6###");
            data[4][2] = new String("0#########1#########2#########3#########4#########5#########6###");
            data[5][2] = new String("0#########1#########2#########3#########4#########5#########6###");    //@@
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
    }  //public class CtTblMdl extends AbstractTableModel


    /**
     *
     *   �O���[�v�ʂ̃��V�s�e�[�u���ꗗ
     *
     */
    class GroupeTable extends JTable {

        private GrTblMdl model = null;

        GroupeTable(){
            super();

            try{
                setName("GroupeTable");
                setBounds(0, 0, 200, 200);
                setAutoCreateColumnsFromModel(true);
                setSelectionMode(ListSelectionModel.SINGLE_SELECTION);

                setLocale(new Locale("ja","JP"));
                setFont(new java.awt.Font("dialog", 0, 12));
                setRowHeight(17);

                model = new GrTblMdl();
                setModel(model);

                DefaultTableColumnModel cmdl = (DefaultTableColumnModel)getColumnModel();
                TableColumn  colum = null;

                // ���V�s�[No
                colum = cmdl.getColumn(0);
                colum.setMaxWidth(40);
                colum.setMinWidth(40);
                colum.setWidth(40);

                // �^�C�g��
                colum = cmdl.getColumn(1);
            }
            catch (Throwable e) {
                CZSystem.handleException(e);
            }
        }

        //
        // ���V�s�I��
        //
        public void valueChanged(ListSelectionEvent e){
            super.valueChanged(e);

//@@            CZSystem.log("CZControlTable","GroupeTable valueChanged [" +
//@@                     getSelectedRow() + "][" + getSelectedColumn() + "]");
        }


        //
        //�O���[�v�ʂ̃��V�s�e�[�u���̃f�[�^��ݒ肷��B
        //@param gr ... �O���[�v,   tbl ... ���V�s�e�[�u��
        public void setData(int gr,int tbl){

//@@            CZSystem.log("CZControlTable","GroupeTable setData [" + gr + "][" + tbl + "]");

            CZSystemCtTitle title   = null;
            String      empty   = new String("");

            for(int i = 0 ; i < 999 ; i++){
                g_table.setValueAt(empty,i,1);
            }

            if( 0 < tbl) {
                for(int i = 0 ; i < table_title.size() ; i++){
                    title = (CZSystemCtTitle)table_title.elementAt(i);
                    if(gr == title.g_no){
                        g_table.setValueAt(title.title,title.r_no-1,1);
                    }
                }
                setRowSelectionInterval(tbl-1,tbl-1);

                Rectangle cellRect = getCellRect(tbl-1,0,false);
                if(cellRect != null){
                    scrollRectToVisible(cellRect);
                }
            } else {
                setRowSelectionInterval(0,0);
                scrollRectToVisible(getCellRect(0,0,false));
            }
            repaint();
        }
    }  //class GroupeTable extends JTable


    /**
     *
     *       ���V�s�e�[�u���N���X�F���f��
     *
     */
    public class GrTblMdl extends AbstractTableModel {

        final   int TBL_COL = 2;
        private int TBL_ROW = 999;

        final String[] names = {" # " ,"�^�C�g��"};

        private Object  data[][];

        GrTblMdl(){
            super();

            data = new Object[TBL_ROW][TBL_COL];

            for(int i = 0 ; i < TBL_ROW ; i++){
                data[i][0] = new Integer(i+1);
                data[i][1] = new String("");
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

    }  //public class GrTblMdl extends AbstractTableModel

    /**
     *
     *   ���ڃe�[�u���ꗗ
     *
     */
    class ValueTable extends JTable {

        private VaTblMdl model = null;

        ValueTable(){
            super();

            try{
                setName("ValueTable");
                setBounds(0, 0, 200, 200);
                setAutoCreateColumnsFromModel(true);
                setSelectionMode(ListSelectionModel.MULTIPLE_INTERVAL_SELECTION);

                setLocale(new Locale("ja","JP"));
                setFont(new java.awt.Font("dialog", 0, 12));
                setRowHeight(17);

                model = new VaTblMdl();
                setModel(model);

                DefaultTableColumnModel cmdl = (DefaultTableColumnModel)getColumnModel();
                TableColumn  colum = null;

                // ����No
                colum = cmdl.getColumn(0);
                colum.setMaxWidth(40);
                colum.setMinWidth(40);
                colum.setWidth(40);

                // ����
                colum = cmdl.getColumn(1);
                colum.setMaxWidth(230);
                colum.setMinWidth(230);
                colum.setWidth(230);

                // �k������
                colum = cmdl.getColumn(2);
                colum.setMaxWidth(100);
                colum.setMinWidth(100);
                colum.setWidth(100);

                // �q������
                colum = cmdl.getColumn(3);
                colum.setMaxWidth(100);
                colum.setMinWidth(100);
                colum.setWidth(100);

                // �o�u�Ή�
                colum = cmdl.getColumn(4);
            }
            catch (Throwable e) {
                CZSystem.handleException(e);
            }
        }

        //
        // ���ڑI��
        //
        public void valueChanged(ListSelectionEvent e){
            super.valueChanged(e);

//@@            CZSystem.log("CZControlTable","ValueTable valueChanged [" +
//@@                         getSelectedRow() + "][" + getSelectedColumn() + "]");

        }

        //
        //  ���ڃf�[�^��ݒ肷��
        //@param ... �O���[�v
        public void setData(int gr){

//@@            CZSystem.log("CZControlTable","ValueTable setData [" + gr + "]");

            CZSystemCtName name =  null;
            String      empty   = "";

            for(int i = 0 ; i < 600 ; i++){
                name = CZSystem.getCtTbName(gr,i+1);

                if(null != name){
                    setValueAt(name.t_name.trim(),i,1);
                    setValueAt(name.l_name.trim(),i,2);
                    setValueAt(name.r_name.trim(),i,3);
                    setValueAt(new Integer(name.pv_no),i,4);
                }
                else {
                    setValueAt(empty,i,1);
                    setValueAt(empty,i,2);
                    setValueAt(empty,i,3);
                    setValueAt(empty,i,4);
                }
            } // for end

            setRowSelectionInterval(0,0);

            Rectangle cellRect = getCellRect(0,0,false);
            if(cellRect != null){
                scrollRectToVisible(cellRect);
            }

            repaint();
        }
    }  //class ValueTable extends JTable


    /**
     *
     *       ���ڃe�[�u���F���f��
     *
     */
    public class VaTblMdl extends AbstractTableModel {

        final   int TBL_COL = 5;
        private int TBL_ROW = 600;

        final String[] names = {" # " ,"����","�k������","�q������","�o�u�Ή�"};

        private Object  data[][];

        VaTblMdl(){
            super();

            data = new Object[TBL_ROW][TBL_COL];

            for(int i = 0 ; i < TBL_ROW ; i++){
                data[i][0] = new Integer(i+1);
                data[i][1] = new String("����");
                data[i][2] = new String("�k������");
                data[i][3] = new String("�q������");
                data[i][4] = new Integer(i+100);
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
    }  //public class VaTblMdl extends AbstractTableModel

    /**
     *
     *   �區�ڃe�[�u���ꗗ
     *
     */
    class T6LagTable extends JTable {

        private T6LagTableMdl model = null;

        T6LagTable(){
            super();

            try{
                setName("T6LagTable");
                setBounds(0, 0, 200, 200);
                setAutoCreateColumnsFromModel(true);
                setSelectionMode(ListSelectionModel.SINGLE_SELECTION);

                setLocale(new Locale("ja","JP"));
                setFont(new java.awt.Font("dialog", 0, 12));
                setRowHeight(17);

                model = new T6LagTableMdl();
                setModel(model);

                DefaultTableColumnModel cmdl = (DefaultTableColumnModel)getColumnModel();
                TableColumn  colum = null;

                // ����No
                colum = cmdl.getColumn(0);
                colum.setMaxWidth(40);
                colum.setMinWidth(40);
                colum.setWidth(40);

                // ����
                colum = cmdl.getColumn(1);
                colum.setMaxWidth(300);
                colum.setMinWidth(300);
                colum.setWidth(300);
            }
            catch (Throwable e) {
                CZSystem.handleException(e);
            }
        }

        //
        // �區�ڑI��
        //
        public void valueChanged(ListSelectionEvent e){
            super.valueChanged(e);

            if(0 > this.getSelectedRow()) return;
            if(0 > c_table.getSelectedRow()) return;
            if(0 > g_table.getSelectedRow()) return;

//@@            CZSystem.log("CZControlTable","T6LagTable ValueTable valueChanged [" +
//@@                         getSelectedRow() + "][" + getSelectedColumn() + "]");

            //�����ڂ����o���B @@@
            int rpNo = ((Integer)g_table.getValueAt(g_table.getSelectedRow(),0)).intValue();
            int lgNo = ((Integer)this.getValueAt(this.getSelectedRow(),0)).intValue();
            t6MidTable_.setData(6,rpNo,lgNo);
            //���ڂ����o���B @@@
            int mdNo = ((Integer)this.getValueAt(0,0)).intValue();
            t6ValTable_.setData(6,lgNo,mdNo);

			/****** 2007.01.24 ADD *****/
            //���ݒl�����o���B
            boolean current;
            if (6 == rpNo) {
                current = true;
            } else {
                current = false;
            }
            t6Current_ = CZSystem.getCtT6Tb(6, rpNo,lgNo, mdNo, current);
            setT6Win_.setDefault( 6, rpNo, lgNo, mdNo );
			/***************************/

        }

        //
        //  �區�ڃf�[�^��ݒ肷��B
        //@param grp ... �O���[�v,rcp ... ���V�s
        public void setData(int grp,int rcp){

//@@            CZSystem.log("CZControlTable","T6LagTable setData [" + grp + "]");

            CZSystemCtT6LagName name =  null;
            String      empty   = "";

            for(int i = 0 ; i < 100 ; i++){
                name = CZSystem.getCtT6LagName(grp, rcp,i+1);

                if(null != name){
                    setValueAt(name.k_name1.trim(),i,1);
                }
                else {
                    setValueAt(empty,i,1);
                }
            } // for end

            setRowSelectionInterval(0,0);

            Rectangle cellRect = getCellRect(0,0,false);
            if(cellRect != null){
                scrollRectToVisible(cellRect);
            }

            repaint();
        }
    } //class T6LagTable extends JTable

    /**
     *
     *       �區�ڃe�[�u���F���f��
     *
     */
    public class T6LagTableMdl extends AbstractTableModel {

        final   int TBL_COL = 2;
        private int TBL_ROW = 100;

        final String[] names = {" # " ,"�區��"};

        private Object  data[][];

        T6LagTableMdl(){
            super();

            data = new Object[TBL_ROW][TBL_COL];

            for(int i = 0 ; i < TBL_ROW ; i++){
                data[i][0] = new Integer(i+1);
                data[i][1] = new String("�區��");
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
    }  //public class T6LagTableMdl extends AbstractTableModel


    /**
     *
     *   �����ڃe�[�u���ꗗ
     *
     */
    class T6MidTable extends JTable {

        private T6MidTableMdl model = null;

        T6MidTable(){
            super();

            try{
                setName("T6LagTable");
                setBounds(0, 0, 200, 200);
                setAutoCreateColumnsFromModel(true);
                setSelectionMode(ListSelectionModel.SINGLE_SELECTION);

                setLocale(new Locale("ja","JP"));
                setFont(new java.awt.Font("dialog", 0, 12));
                setRowHeight(17);

                model = new T6MidTableMdl();
                setModel(model);

                DefaultTableColumnModel cmdl = (DefaultTableColumnModel)getColumnModel();
                TableColumn  colum = null;

                // ����No
                colum = cmdl.getColumn(0);
                colum.setMaxWidth(40);
                colum.setMinWidth(40);
                colum.setWidth(40);

                // ����
                colum = cmdl.getColumn(1);
                colum.setMaxWidth(300);
                colum.setMinWidth(300);
                colum.setWidth(300);
            }
            catch (Throwable e) {
                CZSystem.handleException(e);
            }
        }

        //
        //�@�����ڑI��
        //
        public void valueChanged(ListSelectionEvent e){
            super.valueChanged(e);

            if(0 > this.getSelectedRow()) return;
            if(0 > c_table.getSelectedRow()) return;
            if(0 > g_table.getSelectedRow()) return;
            if(0 > t6LagTable_.getSelectedRow()) return;

//@@            CZSystem.log("CZControlTable","T6MidTable ValueTable valueChanged [" +
//@@                         getSelectedRow() + "][" + getSelectedColumn() + "]");

            //���ڂ����o���B @@@
            int crpNo = ((Integer)c_table.getValueAt(c_table.getSelectedRow(),1)).intValue();
            int rpNo = ((Integer)g_table.getValueAt(g_table.getSelectedRow(),0)).intValue();
            int lgNo = ((Integer)t6LagTable_.getValueAt(t6LagTable_.getSelectedRow(),0)).intValue();
            int mdNo = ((Integer)t6MidTable_.getValueAt(t6MidTable_.getSelectedRow(),0)).intValue();
            t6ValTable_.setData(6,lgNo,mdNo);
            //���ݒl�����o���B
            boolean current;
            if (crpNo == rpNo) {
                current = true;
            } else {
                current = false;
            }
            t6Current_ = CZSystem.getCtT6Tb(6, rpNo,lgNo, mdNo, current);
            setT6Win_.setDefault( 6, rpNo, lgNo, mdNo );

        }

        //
        //  �����ڃe�[�u���̃f�[�^��ݒ肷��B
        //@param grp ... �O���[�v, rcp ... ���V�s
        public void setData(int grp, int rcp, int lag){

//@@            CZSystem.log("CZControlTable","T6MidTable setData [" + grp + ":"+ rcp + ":" + lag + "]");

            CZSystemCtT6MidName name =  null;
            String      empty   = "";

            for(int i = 0 ; i < 100 ; i++){
                name = CZSystem.getCtT6MidName(grp, rcp, lag, i+1);

                if(null != name){
                    setValueAt(name.k_name2.trim(),i,1);
                }
                else {
                    setValueAt(empty,i,1);
                }
            } // for end

            setRowSelectionInterval(0,0);

            Rectangle cellRect = getCellRect(0,0,false);
            if(cellRect != null){
                scrollRectToVisible(cellRect);
            }

            repaint();
        }
    }  //class T6MidTable extends JTable


    /**
     *
     *       �����ڃe�[�u���F���f��
     *
     */
    public class T6MidTableMdl extends AbstractTableModel {

        final   int TBL_COL = 2;
        private int TBL_ROW = 100;

        final String[] names = {" # " ,"������"};

        private Object  data[][];

        T6MidTableMdl(){
            super();

            data = new Object[TBL_ROW][TBL_COL];

            for(int i = 0 ; i < TBL_ROW ; i++){
                data[i][0] = new Integer(i+1);
                data[i][1] = new String("������");
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
    }  //public class T6MidTableMdl extends AbstractTableModel


    /**
     *
     *   T6���ڃe�[�u���ꗗ
     *
     */
    class T6ValTable extends JTable {

        private T6ValTableMdl model = null;

        T6ValTable(){
            super();

            try{
                setName("T6LagTable");
                setBounds(0, 0, 200, 200);
                setAutoCreateColumnsFromModel(true);
                setSelectionMode(ListSelectionModel.SINGLE_SELECTION);

                setLocale(new Locale("ja","JP"));
                setFont(new java.awt.Font("dialog", 0, 12));
                setRowHeight(17);

                model = new T6ValTableMdl();
                setModel(model);

                DefaultTableColumnModel cmdl = (DefaultTableColumnModel)getColumnModel();
                TableColumn  colum = null;

                // ����No
                colum = cmdl.getColumn(0);
                colum.setMaxWidth(40);
                colum.setMinWidth(40);
                colum.setWidth(40);

                // ����
                colum = cmdl.getColumn(1);
                colum.setMaxWidth(300);
                colum.setMinWidth(300);
                colum.setWidth(300);
            }
            catch (Throwable e) {
                CZSystem.handleException(e);
            }
        }

        //
        // �s�U���ڑI��
        //
        public void valueChanged(ListSelectionEvent e){
            super.valueChanged(e);

//@@            CZSystem.log("CZControlTable","T6ValTable ValueTable valueChanged [" +
//@@                         getSelectedRow() + "][" + getSelectedColumn() + "]");

        }

        //
        //
        //
        public void setData(int gr, int lg, int md){

//@@            CZSystem.log("CZControlTable","T6ValTable setData [" + gr + "][" + lg + "][" + md + "]");

            CZSystemCtT6Name name =  null;
            String      empty   = "";

            for(int i = 0 ; i < 600 ; i++){
                name = CZSystem.getCtT6Name(gr, lg, md, i+1);

                if(null != name){
                    setValueAt(name.k_name.trim(),i,1);
                }
                else {
                    setValueAt(empty,i,1);
                }
            } // for end

            setRowSelectionInterval(0,0);

            Rectangle cellRect = getCellRect(0,0,false);
            if(cellRect != null){
                scrollRectToVisible(cellRect);
            }

            repaint();
        }
    }  //class T6ValTable extends JTable


    /**
     *
     *       T6���ڃe�[�u���F���f��
     *
     */
    public class T6ValTableMdl extends AbstractTableModel {

        final   int TBL_COL = 2;
        private int TBL_ROW = 600;

        final String[] names = {" # " ,"����"};

        private Object  data[][];

        T6ValTableMdl(){
            super();

            data = new Object[TBL_ROW][TBL_COL];

            for(int i = 0 ; i < TBL_ROW ; i++){
                data[i][0] = new Integer(i+1);
                data[i][1] = new String("����");
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
    } //public class T6ValTableMdl extends AbstractTableModel

    /**
     *
     *       �����W�pWindow
     *
     */
    public class LimitWin extends JDialog {

        private CZSystemCtName  ct_name     = null;

        private ItemText    item_name   = null;
        private JComboBox   sort_kubun  = null;
        private PVText      pv_no       = null;

        private ItemText    l_name      = null;
        private MinMaxText  l_min       = null;
        private MinMaxText  l_max       = null;
        private DigitText   l_digit     = null;
        private UnitText    l_unit      = null;

        private ItemText    r_name      = null;
        private MinMaxText  r_min       = null;
        private MinMaxText  r_max       = null;
        private DigitText   r_digit     = null;
        private UnitText    r_unit      = null;

        private TText       op_name     = null;

        private JButton     unit_send_button   = null;
        private JButton     unit_cancel_button = null;


        private String      sendOp      = null;
        private String      sendName    = null;
        private int         sendPVNo    = 1;
        private int         sendSort    = 1;

        private String      sendLName   = null;
        private String      sendRName   = null;
        private String      sendLUnit   = null;
        private String      sendRUnit   = null;
        private float       sendLMin    = 0.0f;
        private float       sendLMax    = 1.0f;
        private float       sendRMin    = 0.0f;
        private float       sendRMax    = 1.0f;
        private int         sendLDigit  = 0;
        private int         sendRDigit  = 0;


        //
        //�R���X�g���N�^
        //
        LimitWin(){
            super();

            setTitle("����e�[�u�����ڐݒ�");
            setSize(845,310);
            setResizable(false);
            setModal(true);

            getContentPane().setLayout(null);
            // ����n�Q�Ƌ@�\    @20131021
            if(CZSystemDefine.REFERENCE_RUN == CZSystem.getRunLevel()){
                getContentPane().setBackground(CZSystemDefine.DEFAULT_REFERENCE_BACKGROUND_COL);
            }else{
                getContentPane().setBackground(CZSystemDefine.DEFAULT_BACKGROUND_COL);
            }

            JLabel lab = null;

            lab = new JLabel("��                ��",JLabel.CENTER);
            lab.setBounds(20+80, 20, 300, 24);
            lab.setLocale(new Locale("ja","JP"));
            lab.setFont(new java.awt.Font("dialog", 0, 16));
            lab.setBorder(new Flush3DBorder());
            lab.setForeground(java.awt.Color.black);
            getContentPane().add(lab);

            lab = new JLabel("�\�[�g�敪",JLabel.CENTER);
            lab.setBounds(330+80, 20, 100, 24);
            lab.setLocale(new Locale("ja","JP"));
            lab.setFont(new java.awt.Font("dialog", 0, 16));
            lab.setBorder(new Flush3DBorder());
            lab.setForeground(java.awt.Color.black);
            getContentPane().add(lab);

            lab = new JLabel("�o�u",JLabel.CENTER);
            lab.setBounds(440+80, 20, 50, 24);
            lab.setLocale(new Locale("ja","JP"));
            lab.setFont(new java.awt.Font("dialog", 0, 16));
            lab.setBorder(new Flush3DBorder());
            lab.setForeground(java.awt.Color.black);
            getContentPane().add(lab);

            lab = new JLabel("��                ��",JLabel.CENTER);
            lab.setBounds(100, 90, 300, 24);
            lab.setLocale(new Locale("ja","JP"));
            lab.setFont(new java.awt.Font("dialog", 0, 16));
            lab.setBorder(new Flush3DBorder());
            lab.setForeground(java.awt.Color.black);
            getContentPane().add(lab);

            lab = new JLabel("�l����",JLabel.CENTER);
            lab.setBounds(410, 90, 100, 24);
            lab.setLocale(new Locale("ja","JP"));
            lab.setFont(new java.awt.Font("dialog", 0, 16));
            lab.setBorder(new Flush3DBorder());
            lab.setForeground(java.awt.Color.black);
            getContentPane().add(lab);

            lab = new JLabel("�l����",JLabel.CENTER);
            lab.setBounds(520, 90, 100, 24);
            lab.setLocale(new Locale("ja","JP"));
            lab.setFont(new java.awt.Font("dialog", 0, 16));
            lab.setBorder(new Flush3DBorder());
            lab.setForeground(java.awt.Color.black);
            getContentPane().add(lab);

            lab = new JLabel("��",JLabel.CENTER);
            lab.setBounds(630, 90, 25, 24);
            lab.setLocale(new Locale("ja","JP"));
            lab.setFont(new java.awt.Font("dialog", 0, 16));
            lab.setBorder(new Flush3DBorder());
            lab.setForeground(java.awt.Color.black);
            getContentPane().add(lab);

            lab = new JLabel("�P        ��",JLabel.CENTER);
            lab.setBounds(665, 90, 150, 24);
            lab.setLocale(new Locale("ja","JP"));
            lab.setFont(new java.awt.Font("dialog", 0, 16));
            lab.setBorder(new Flush3DBorder());
            lab.setForeground(java.awt.Color.black);
            getContentPane().add(lab);

            lab = new JLabel("�k��",JLabel.CENTER);
            lab.setBounds(20, 114, 80, 24);
            lab.setLocale(new Locale("ja","JP"));
            lab.setFont(new java.awt.Font("dialog", 0, 16));
            lab.setBorder(new Flush3DBorder());
            lab.setForeground(java.awt.Color.black);
            getContentPane().add(lab);

            lab = new JLabel("�q��",JLabel.CENTER);
            lab.setBounds(20, 138, 80, 24);
            lab.setLocale(new Locale("ja","JP"));
            lab.setFont(new java.awt.Font("dialog", 0, 16));
            lab.setBorder(new Flush3DBorder());
            lab.setForeground(java.awt.Color.black);
            getContentPane().add(lab);

            lab = new JLabel("�ݒ��",JLabel.CENTER);
            lab.setBounds(20, 180, 100, 24);
            lab.setLocale(new Locale("ja","JP"));
            lab.setFont(new java.awt.Font("dialog", 0, 16));
            lab.setBorder(new Flush3DBorder());
            lab.setForeground(java.awt.Color.black);
            getContentPane().add(lab);

            item_name = new ItemText();
            item_name.setBounds(20+80, 44, 300, 24);
            item_name.setLocale(new Locale("ja","JP"));
            item_name.setFont(new java.awt.Font("dialog", 0, 16));
            item_name.setBorder(new Flush3DBorder());
            item_name.setForeground(java.awt.Color.black);
            getContentPane().add(item_name);

            sort_kubun = new JComboBox();
            sort_kubun.setBounds(330+80, 44, 100, 24);
            sort_kubun.setLocale(new Locale("ja","JP"));
            sort_kubun.setFont(new java.awt.Font("dialog", 0, 16));
            sort_kubun.setForeground(java.awt.Color.black);
            sort_kubun.addItem("����");
            sort_kubun.addItem("�~��");
            sort_kubun.setFocusable(false);	/* 2007.08.22 */
            getContentPane().add(sort_kubun);

            pv_no = new PVText();
            pv_no.setBounds(440+80, 44, 50, 24);
            pv_no.setLocale(new Locale("ja","JP"));
            pv_no.setFont(new java.awt.Font("dialog", 0, 16));
            pv_no.setBorder(new Flush3DBorder());
            pv_no.setForeground(java.awt.Color.black);
            getContentPane().add(pv_no);

            // �k��
            l_name = new ItemText();
            l_name.setBounds(100, 114, 300, 24);
            l_name.setLocale(new Locale("ja","JP"));
            l_name.setFont(new java.awt.Font("dialog", 0, 16));
            l_name.setBorder(new Flush3DBorder());
            l_name.setForeground(java.awt.Color.black);
            getContentPane().add(l_name);

            l_min = new MinMaxText();
            l_min.setBounds(410, 114, 100, 24);
            l_min.setLocale(new Locale("ja","JP"));
            l_min.setFont(new java.awt.Font("dialog", 0, 16));
            l_min.setBorder(new Flush3DBorder());
            l_min.setForeground(java.awt.Color.black);
            getContentPane().add(l_min);

            l_max = new MinMaxText();
            l_max.setBounds(520, 114, 100, 24);
            l_max.setLocale(new Locale("ja","JP"));
            l_max.setFont(new java.awt.Font("dialog", 0, 16));
            l_max.setBorder(new Flush3DBorder());
            l_max.setForeground(java.awt.Color.black);
            getContentPane().add(l_max);

            l_digit = new DigitText();
            l_digit.setBounds(630, 114, 25, 24);
            l_digit.setLocale(new Locale("ja","JP"));
            l_digit.setFont(new java.awt.Font("dialog", 0, 16));
            l_digit.setBorder(new Flush3DBorder());
            l_digit.setForeground(java.awt.Color.black);
            getContentPane().add(l_digit);

            l_unit = new UnitText();
            l_unit.setBounds(665, 114, 150, 24);
            l_unit.setLocale(new Locale("ja","JP"));
            l_unit.setFont(new java.awt.Font("dialog", 0, 16));
            l_unit.setBorder(new Flush3DBorder());
            l_unit.setForeground(java.awt.Color.black);
            getContentPane().add(l_unit);

            // �q��
            r_name = new ItemText();
            r_name.setBounds(100, 138, 300, 24);
            r_name.setLocale(new Locale("ja","JP"));
            r_name.setFont(new java.awt.Font("dialog", 0, 16));
            r_name.setBorder(new Flush3DBorder());
            r_name.setForeground(java.awt.Color.black);
            getContentPane().add(r_name);

            r_min = new MinMaxText();
            r_min.setBounds(410, 138, 100, 24);
            r_min.setLocale(new Locale("ja","JP"));
            r_min.setFont(new java.awt.Font("dialog", 0, 16));
            r_min.setBorder(new Flush3DBorder());
            r_min.setForeground(java.awt.Color.black);
            getContentPane().add(r_min);

            r_max = new MinMaxText();
            r_max.setBounds(520, 138, 100, 24);
            r_max.setLocale(new Locale("ja","JP"));
            r_max.setFont(new java.awt.Font("dialog", 0, 16));
            r_max.setBorder(new Flush3DBorder());
            r_max.setForeground(java.awt.Color.black);
            getContentPane().add(r_max);

            r_digit = new DigitText();
            r_digit.setBounds(630, 138, 25, 24);
            r_digit.setLocale(new Locale("ja","JP"));
            r_digit.setFont(new java.awt.Font("dialog", 0, 16));
            r_digit.setBorder(new Flush3DBorder());
            r_digit.setForeground(java.awt.Color.black);
            getContentPane().add(r_digit);

            r_unit = new UnitText();
            r_unit.setBounds(665, 138, 150, 24);
            r_unit.setLocale(new Locale("ja","JP"));
            r_unit.setFont(new java.awt.Font("dialog", 0, 16));
            r_unit.setBorder(new Flush3DBorder());
            r_unit.setForeground(java.awt.Color.black);
            getContentPane().add(r_unit);

            // �I�y���[�^��
            op_name = new TText();
            op_name.setBounds(120, 180, 140, 24);
            getContentPane().add(op_name);

            unit_send_button = new JButton();
            unit_send_button = new JButton("��  �s");
            unit_send_button.setBounds(260, 180, 100, 24);
            unit_send_button.setLocale(new Locale("ja","JP"));
            unit_send_button.setFont(new java.awt.Font("dialog", 0, 18));
            unit_send_button.setBorder(new Flush3DBorder());
            unit_send_button.setForeground(java.awt.Color.black);
            unit_send_button.addActionListener(new UnitSendButton());
            getContentPane().add(unit_send_button);

            unit_cancel_button = new JButton("�I  ��");
            unit_cancel_button.setBounds(715, 180, 100, 24);
            unit_cancel_button.setLocale(new Locale("ja","JP"));
            unit_cancel_button.setFont(new java.awt.Font("dialog", 0, 18));
            unit_cancel_button.setBorder(new Flush3DBorder());
            unit_cancel_button.setForeground(java.awt.Color.black);
            unit_cancel_button.addActionListener(new UnitCancelButton());
            getContentPane().add(unit_cancel_button);

        }

        //
        //
        //
        public boolean setDefault(CZSystemCtName _name){
            if(null == _name) return false;

            ct_name = _name;

            item_name.setText(ct_name.t_name);

            if(2 == ct_name.k_sort){
                sort_kubun.setSelectedIndex(1);
            }
            else {
                sort_kubun.setSelectedIndex(0);
            }

            pv_no.setText(Integer.toString(ct_name.pv_no));

            l_name.setText(ct_name.l_name.trim());
            l_min.setText(Float.toString(ct_name.l_min));
            l_max.setText(Float.toString(ct_name.l_max));
            l_digit.setText(Integer.toString(ct_name.l_keta));
            l_unit.setText(ct_name.l_unit.trim());

            r_name.setText(ct_name.r_name.trim());
            r_min.setText(Float.toString(ct_name.r_min));
            r_max.setText(Float.toString(ct_name.r_max));
            r_digit.setText(Integer.toString(ct_name.r_keta));
            r_unit.setText(ct_name.r_unit.trim());

            op_name.setText("");
            

            return true;
        }

        //
        //
        //
        private boolean setUnitSendStatus(){
            sendOp = op_name.getText();
            if(1 > sendOp.length()){
                return false;
            }

            sendName = item_name.getText();
            if(1 > sendName.length()){
                return false;
            }

            switch(sort_kubun.getSelectedIndex()){
                case 1  : sendSort = 2;
                      break;

                default : sendSort = 1;
                      break;
            }


            sendLName = l_name.getText();
            if(1 > sendLName.length()){
                return false;
            }

            sendRName = r_name.getText();
            if(1 > sendRName.length()){
                return false;
            }

            sendLUnit = l_unit.getText();
            if(1 > sendLUnit.length()){
                return false;
            }

            sendRUnit = r_unit.getText();
            if(1 > sendRUnit.length()){
                return false;
            }


            try{
                sendPVNo   = Integer.parseInt(pv_no.getText());

                sendLMin   = Float.parseFloat(l_min.getText());
                sendLMax   = Float.parseFloat(l_max.getText());
                sendLDigit = Integer.parseInt(l_digit.getText());

                sendRMin   = Float.parseFloat(r_min.getText());
                sendRMax   = Float.parseFloat(r_max.getText());
                sendRDigit = Integer.parseInt(r_digit.getText());
            }
            catch (Exception e){
                return false;
            }

            if(sendLMin >= sendLMax) return false;
            if(sendRMin >= sendRMax) return false;
            if(0 > sendLDigit) return false;
            if(0 > sendRDigit) return false;

            return true;
        }

        /**
         *       ���ږ�����͂���TextField
         */
        public class ItemText extends JTextField {

            ItemText(){
                super();
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

                //
                //
                public void insertString( int offset, String str, AttributeSet a )
                                                throws BadLocationException {

                    String tmp = new String(getText(0,getLength()) + str);
                    byte   b[];

                    try{
                        b = tmp.getBytes("SJIS");
                    }
                    catch(Exception e){
                        CZSystem.log("CZControlTable","LimitWin ItemText [" + e + "]");
                        return;
                    }

//@@                    CZSystem.log("CZControlTable","LimitWin ItemText [" + tmp + "][" + b + "][" + b.length + "]");

//@@@                    if(32 < b.length) return;
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


        /**
         *       �Ή��o�u����͂���TextField
         */
        public class PVText extends JTextField {

            PVText(){
                super();
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

                    if(2 < getLength()) return;
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

        /**
         *       �l�����l��������͂���TextField
         */
        public class MinMaxText extends JTextField {

            MinMaxText(){
                super();
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
                String validValues = "0123456789.-";

                //
                //
                public void insertString( int offset, String str, AttributeSet a )
                                    throws BadLocationException {

                    if(9 < getLength()) return;
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

        /**
         *       ������͂���TextField
         */
        public class DigitText extends JTextField {

            DigitText(){
                super();
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
                String validValues = "0123456";

                //
                //
                public void insertString( int offset, String str, AttributeSet a )
                                        throws BadLocationException {

                    if(0 < getLength()) return;
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

        /**
         *       �P�ʂ���͂���TextField
         */
        public class UnitText extends JTextField {

            UnitText(){
                super();
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

                //
                //
                public void insertString( int offset, String str, AttributeSet a )
                                                throws BadLocationException {

                    String tmp = new String(getText(0,getLength()) + str);
                    byte   b[];

                    try{
                        b = tmp.getBytes("SJIS");
                    }
                    catch(Exception e){
                        CZSystem.log("CZControlTable","LimitWin ItemText [" + e + "]");
                        return;
                    }

//@@                    CZSystem.log("CZControlTable","LimitWin ItemText [" + tmp + "][" + b + "][" + b.length + "]");

                    if(16 < b.length) return;
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

        /**
         *       �ݒ�҂���͂���TextField
        */
        public class TText extends JTextField {

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

        /**
         *
         *      ���s�{�^��
         *
         */
        class UnitSendButton implements ActionListener {
            public void actionPerformed(ActionEvent ev){

                if(!setUnitSendStatus()){
                    Object msg[] = {"����e�[�u�����ڍX�V",
                                    "�ݒ�ҁA���ځAMin�AMax�A����",
                                    "�������Ă�������"};
                    errorMsg(msg);
                    return;
                }
/*
                CZSystem.log("CZControlTable","LimitWin UnitSendButton-->[" +
                                                        ct_name.g_no    + "][" + ct_name.t_no + "][" +
                                                        sendOp     + "][" + sendName  + "][" +
                                                        sendPVNo   + "][" + sendSort  + "][" +
                                                        sendLName  + "][" + sendRName + "][" +
                                                        sendLUnit  + "][" + sendRUnit + "][" +
                                                        sendLMin   + "][" + sendLMax  + "][" +
                                                        sendRMin   + "][" + sendRMax  + "][" +
                                                        sendLDigit + "][" + sendRDigit+ "]");
*/
                CZSystem.log("CZControlTable","LimitWin UnitSendButton-->[" + sendName + "]");

                CZParamControlDefine s = new CZParamControlDefine();
                s.setTname(sendName);
                s.setLname(sendLName);
                s.setLtani(sendLUnit);
                s.setLmin(sendLMin);
                s.setLmax(sendLMax);
                s.setLpoint(sendLDigit);
                s.setRname(sendRName);
                s.setRtani(sendRUnit);
                s.setRmin(sendRMin);
                s.setRmax(sendRMax);
                s.setRpoint(sendRDigit);
                s.setSort(sendSort);
                s.setPvno(sendPVNo);

                //Send
                if(!CZSystem.CZControlDefineExchange(sendOp, ct_name.g_no , ct_name.t_no, s)){

                    Object msg[] = {"����e�[�u�����ڍX�V",
                                "�X�V�����s���܂���",
                                "�Ǘ��҂ɖ₢���킹�Ă�������"};
                    errorMsg(msg);
                    return;
                }

                return ;
            }
        }


        /**
         *      �I���{�^��
        */
        class UnitCancelButton implements ActionListener {
            public void actionPerformed(ActionEvent ev){
		        if( 0 != CZSystemDefine.TIMER_FLG ){
					at.stop();
					t.restart();
					CZSystem.log("CZControlTable","�A���[����ʃI�[�v�����X�^�[�g�i�����W�I���j");
				}
                setVisible(false);
            }
        }
    } //public class LimitWin extends JDialog

    /**
     *
     *       ���V�s�[�^�C�g���ݒ�pWindow
     *
     */
    public class TitleWin extends JDialog {

        private int     current_group   = 0;
        private int     current_recip   = 0;

        private JLabel      group_name  = null;
        private JLabel      recip_no    = null;

        private ItemText    item_name   = null;

        private TText       op_name     = null;

        private JButton     title_send_button   = null;
        private JButton     title_cancel_button = null;

        private String      sendOp      = null;
        private String      sendTitle   = null;

        //
        //�R���X�g���N�^
        //
        TitleWin(){
            super();

            setTitle("����e�[�u���^�C�g���ݒ�");
            setSize(670,160);
            setResizable(false);
            setModal(true);

            getContentPane().setLayout(null);
            // ����n�Q�Ƌ@�\    @20131021
            if(CZSystemDefine.REFERENCE_RUN == CZSystem.getRunLevel()){
                getContentPane().setBackground(CZSystemDefine.DEFAULT_REFERENCE_BACKGROUND_COL);
            }else{
                getContentPane().setBackground(CZSystemDefine.DEFAULT_BACKGROUND_COL);
            }


            group_name = new JLabel("�O���[�v",JLabel.CENTER);
            group_name.setBounds(20, 20, 80, 24);
            group_name.setLocale(new Locale("ja","JP"));
            group_name.setFont(new java.awt.Font("dialog", 0, 16));
            group_name.setBorder(new Flush3DBorder());
            group_name.setForeground(java.awt.Color.black);
            getContentPane().add(group_name);

            recip_no = new JLabel("���V�s",JLabel.CENTER);
            recip_no.setBounds(20, 44, 80, 24);
            recip_no.setLocale(new Locale("ja","JP"));
            recip_no.setFont(new java.awt.Font("dialog", 0, 16));
            recip_no.setBorder(new Flush3DBorder());
            recip_no.setForeground(java.awt.Color.black);
            getContentPane().add(recip_no);

            JLabel lab = null;

            lab = new JLabel("�^            �C            �g            ��",JLabel.CENTER);
            lab.setBounds(20+80, 20, 540, 24);
            lab.setLocale(new Locale("ja","JP"));
            lab.setFont(new java.awt.Font("dialog", 0, 16));
            lab.setBorder(new Flush3DBorder());
            lab.setForeground(java.awt.Color.black);
            getContentPane().add(lab);

            item_name = new ItemText();
            item_name.setBounds(20+80, 44, 540, 24);
            item_name.setLocale(new Locale("ja","JP"));
            item_name.setFont(new java.awt.Font("dialog", 0, 16));
            item_name.setBorder(new Flush3DBorder());
            item_name.setForeground(java.awt.Color.black);
            getContentPane().add(item_name);

            // �I�y���[�^��
            lab = new JLabel("�ݒ��",JLabel.CENTER);
            lab.setBounds(20, 92, 100, 24);
            lab.setLocale(new Locale("ja","JP"));
            lab.setFont(new java.awt.Font("dialog", 0, 16));
            lab.setBorder(new Flush3DBorder());
            lab.setForeground(java.awt.Color.black);
            getContentPane().add(lab);

            op_name = new TText();
            op_name.setBounds(120, 92, 140, 24);
            getContentPane().add(op_name);

            title_send_button = new JButton();
            title_send_button = new JButton("��  �s");
            title_send_button.setBounds(260, 92, 100, 24);
            title_send_button.setLocale(new Locale("ja","JP"));
            title_send_button.setFont(new java.awt.Font("dialog", 0, 18));
            title_send_button.setBorder(new Flush3DBorder());
            title_send_button.setForeground(java.awt.Color.black);
            title_send_button.addActionListener(new TitleSendButton());
            getContentPane().add(title_send_button);

            title_cancel_button = new JButton("�I  ��");
            title_cancel_button.setBounds(540, 92, 100, 24);
            title_cancel_button.setLocale(new Locale("ja","JP"));
            title_cancel_button.setFont(new java.awt.Font("dialog", 0, 18));
            title_cancel_button.setBorder(new Flush3DBorder());
            title_cancel_button.setForeground(java.awt.Color.black);
            title_cancel_button.addActionListener(new TitleCancelButton());
            getContentPane().add(title_cancel_button);

        }


        //
        //
        //
        public boolean setDefault(int grp,String _grp,int number,String _ttl){
            current_group = grp;
            current_recip = number;

//@@            CZSystem.log("CZControlTable","TitleWin setDefault [" + _ttl + "]");
            group_name.setText(_grp);
            recip_no.setText(new String("[" + number + "]"));
			/* 2007.04.18 y.k �����R�[�h�΍� */
            item_name.setText(_ttl.trim());
            op_name.setText("");
            
            return true;
        }

        //
        //
        //
        public boolean setTitleSendStatus(){

            if(T1 > current_group) return false;
//            if(T5 < current_group) return false;
            if(T6 < current_group) return false;		/* ����e�[�u���^�C�g���i�s�U�Ή��j 2004.03.16 */
            if(1  > current_recip) return false;
            if(999 < current_recip) return false;

            sendOp = op_name.getText();
            if(1 > sendOp.length()){
                return false;
            }

            sendTitle = item_name.getText();
            if(1 > sendTitle.length()){
                return false;
            }
            return true;
        }

        /**
         *       ����e�[�u���F�^�C�g������͂���TextField
         */
        public class ItemText extends JTextField {

            ItemText(){
                    super();
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

                //
                //
                public void insertString( int offset, String str, AttributeSet a )
                                                        throws BadLocationException {

                    String tmp = new String(getText(0,getLength()) + str);
                    byte   b[];

                    try{
                        b = tmp.getBytes("SJIS");
                    }
                    catch(Exception e){
                        CZSystem.log("CZControlTable","TitleWin ItemText [" + e + "]");
                        return;
                    }

//@@                    CZSystem.log("CZControlTable","TitleWin ItemText [" + tmp + "][" + b + "][" + b.length + "]");

                    if(64 < b.length) return;
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

        /**
         *       �ݒ�҂���͂���TextField
         */
        public class TText extends JTextField {

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

        /**
         *      ���s�{�^��
         */
        class TitleSendButton implements ActionListener {

            public void actionPerformed(ActionEvent ev){

                if(!setTitleSendStatus()){
                    Object msg[] = {"����e�[�u���^�C�g���ύX",
                                    "�O���[�v�A���V�s�[�A�ݒ�ҁA�^�C�g����",
                                    "�������Ă�������"};
                    errorMsg(msg);
                    return;
                }

                //Send
//@@                CZSystem.log("CZControlTable","TitleWin TitleSendButton-->[" +
//@@                                sendOp + "][" + current_group + "][" +
//@@                                current_recip + "][" + sendTitle + "]");

                if(!CZSystem.CZControlTitleExchange(sendOp,current_group,current_recip,sendTitle)){

                    Object msg[] = {"����e�[�u���^�C�g���ύX",
                                    "�ύX�����s���܂���",
                                    "�Ǘ��҂ɖ₢���킹�Ă�������"};
                    errorMsg(msg);
                    return;
                }
                return ;
            }
        }

        /**
         *      �I���{�^��
         */
        class TitleCancelButton implements ActionListener {
            public void actionPerformed(ActionEvent ev){
				if( 0 != CZSystemDefine.TIMER_FLG ){
					at.stop();
					t.restart();
					CZSystem.log("CZControlTable","�A���[����ʃI�[�v�����X�^�[�g�i�^�C�g���I���j");
				}
                setVisible(false);
            }
        }
    }  //public class TitleWin extends JDialog

    /*
    *       T6�區�ڐݒ�pWindow
    */
    public class T6LagSetWin extends JDialog {

        private int     current_group   = 0;
        private int     current_recip   = 0;
        private int     current_lag     = 0;

        private JLabel      group_name  = null;
        private JLabel      recip_name  = null;
        private JLabel      lagNo       = null;

        private ItemText    item_name   = null;

        private TText       op_name     = null;

        private JButton     lagSetWinSendButton   = null;
        private JButton     lagSetWinCancelButton = null;

        private String      sendOp      = null;
        private String      sendLagName = null;

        //
        //
        //
        T6LagSetWin(){
            super();

            setTitle("�s�U�區�ږ��ݒ�");
            setSize(670,160);
            setResizable(false);
            setModal(true);

            getContentPane().setLayout(null);
            // ����n�Q�Ƌ@�\    @20131021
            if(CZSystemDefine.REFERENCE_RUN == CZSystem.getRunLevel()){
                getContentPane().setBackground(CZSystemDefine.DEFAULT_REFERENCE_BACKGROUND_COL);
            }else{
                getContentPane().setBackground(CZSystemDefine.DEFAULT_BACKGROUND_COL);
            }


            group_name = new JLabel("�O���[�v",JLabel.CENTER);
            group_name.setBounds(20, 20, 80, 24);
            group_name.setLocale(new Locale("ja","JP"));
            group_name.setFont(new java.awt.Font("dialog", 0, 16));
            group_name.setBorder(new Flush3DBorder());
            group_name.setForeground(java.awt.Color.black);
            getContentPane().add(group_name);

            recip_name = new JLabel("���V�s",JLabel.CENTER);
            recip_name.setBounds(20+80, 20, 540, 24);
            recip_name.setLocale(new Locale("ja","JP"));
            recip_name.setFont(new java.awt.Font("dialog", 0, 16));
            recip_name.setBorder(new Flush3DBorder());
            recip_name.setForeground(java.awt.Color.black);
            getContentPane().add(recip_name);

            lagNo = new JLabel("����N0",JLabel.CENTER);
            lagNo.setBounds(20, 44, 80, 24);
            lagNo.setLocale(new Locale("ja","JP"));
            lagNo.setFont(new java.awt.Font("dialog", 0, 16));
            lagNo.setBorder(new Flush3DBorder());
            lagNo.setForeground(java.awt.Color.black);
            getContentPane().add(lagNo);

            item_name = new ItemText();
            item_name.setBounds(20+80, 44, 540, 24);
            item_name.setLocale(new Locale("ja","JP"));
            item_name.setFont(new java.awt.Font("dialog", 0, 16));
            item_name.setBorder(new Flush3DBorder());
            item_name.setForeground(java.awt.Color.black);
            getContentPane().add(item_name);

            JLabel lab = null;

            // �I�y���[�^��
            lab = new JLabel("�ݒ��",JLabel.CENTER);
            lab.setBounds(20, 92, 100, 24);
            lab.setLocale(new Locale("ja","JP"));
            lab.setFont(new java.awt.Font("dialog", 0, 16));
            lab.setBorder(new Flush3DBorder());
            lab.setForeground(java.awt.Color.black);
            getContentPane().add(lab);

            op_name = new TText();
            op_name.setBounds(120, 92, 140, 24);
            getContentPane().add(op_name);

            lagSetWinSendButton = new JButton("��  �s");
            lagSetWinSendButton.setBounds(260, 92, 100, 24);
            lagSetWinSendButton.setLocale(new Locale("ja","JP"));
            lagSetWinSendButton.setFont(new java.awt.Font("dialog", 0, 18));
            lagSetWinSendButton.setBorder(new Flush3DBorder());
            lagSetWinSendButton.setForeground(java.awt.Color.black);
            lagSetWinSendButton.addActionListener(new LagSetWinSendButton());
            getContentPane().add(lagSetWinSendButton);

            lagSetWinCancelButton = new JButton("�I  ��");
            lagSetWinCancelButton.setBounds(540, 92, 100, 24);
            lagSetWinCancelButton.setLocale(new Locale("ja","JP"));
            lagSetWinCancelButton.setFont(new java.awt.Font("dialog", 0, 18));
            lagSetWinCancelButton.setBorder(new Flush3DBorder());
            lagSetWinCancelButton.setForeground(java.awt.Color.black);
            lagSetWinCancelButton.addActionListener(new LagSetWinCancelButton());
            getContentPane().add(lagSetWinCancelButton);

        }


        //
        //
        //
        public boolean setDefault(int grp,String _grp,int rcp,String _rcp,int lag, String _lagName){

//@@            CZSystem.log("CZControlTable","T6LagSetWin setDefault() [" + _lagName + "]");

            current_group = grp;
            current_recip = rcp;
            current_lag   = lag;

            group_name.setText( _grp );
            recip_name.setText( _rcp );
            lagNo.setText(new String("" + lag + ""));
            item_name.setText(_lagName);
            op_name.setText("");
            return true;
        }

        //
        //
        //
        public boolean setLagSendStatus(){

            if(T6 != current_group) return false;
            if(1  > current_recip) return false;
            if(999 < current_recip) return false;

            sendOp = op_name.getText();
            if(1 > sendOp.length()){
                return false;
            }

            sendLagName = item_name.getText();
            if(1 > sendLagName.length()){
                return false;
            }
            return true;
        }

        /*
        *       ����e�[�u���F�^�C�g������͂���TextField
        */
        public class ItemText extends JTextField {

            ItemText(){
                    super();
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

                //
                //
                public void insertString( int offset, String str, AttributeSet a )
                                                        throws BadLocationException {

                    String tmp = new String(getText(0,getLength()) + str);
                    byte   b[];

                    try{
                        b = tmp.getBytes("SJIS");
                    }
                    catch(Exception e){
                        CZSystem.log("CZControlTable","T6LagSetWin ItemText [" + e + "]");
                        return;
                    }

//@@                    CZSystem.log("CZControlTable","T6LagSetWin ItemText [" + tmp + "][" + b + "][" + b.length + "]");

                    if(64 < b.length) return;
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

        /*
        *
        *       �ݒ�҂���͂���TextField
        *
        */
        public class TText extends JTextField {

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

        /*
        *
        *
        *
        */

        class LagSetWinSendButton implements ActionListener {

            public void actionPerformed(ActionEvent ev){

                if(!setLagSendStatus()){
                    Object msg[] = {"����e�[�u��:�s�U�區�ڕύX",
                                    "�O���[�v�A���V�s�[�A�ݒ�ҁA�^�C�g����",
                                    "�������Ă�������"};
                    errorMsg(msg);
                    return;
                }

                //Send
//@@                CZSystem.log("CZControlTable","T6LagSetWin LagSetWinSendButton-->[" +
//@@                                sendOp + "][" + current_group + "][" +
//@@                                current_recip + "][" + sendLagName + "]");

/*@@
                if(!CZSystem.CZControlT6LagExchange(sendOp,current_group,current_recip,current_lag,sendLagName)){

                    Object msg[] = {"����e�[�u��:�s�U�區�ڕύX",
                                    "�ύX�����s���܂���",
                                    "�Ǘ��҂ɖ₢���킹�Ă�������"};
                    errorMsg(msg);
                    return;
                }
@@*/
                return ;
            }
        }


        /*
        *
        *
        *
        */
        class LagSetWinCancelButton implements ActionListener {
            public void actionPerformed(ActionEvent ev){
                setVisible(false);
            }
        }
    }  //public class T6LagSetWin extends JDialog

    /*
    *       T6�����ڐݒ�pWindow
    */
    public class T6MidSetWin extends JDialog {

        private int     current_group   = 0;
        private int     current_recip   = 0;
        private int     current_large   = 0;
        private int     current_midle   = 0;

        private JLabel      group_name  = null;
        private JLabel      recip_name  = null;
        private JLabel      large_name  = null;
        private JLabel      large_no    = null;
        private JLabel      midle_no    = null;

        private ItemText    midName     = null;

        private TText       op_name     = null;

        private JButton     midSetWinSendButton   = null;
        private JButton     midSetWinCancelButton = null;

        private String      sendOp      = null;
        private String      sendMidName = null;

        //
        //
        //
        T6MidSetWin(){
            super();

            setTitle("�s�U�����ږ��ݒ�");
            setSize(680,200);
            setResizable(false);
            setModal(true);

            getContentPane().setLayout(null);
            // ����n�Q�Ƌ@�\    @20131021
            if(CZSystemDefine.REFERENCE_RUN == CZSystem.getRunLevel()){
                getContentPane().setBackground(CZSystemDefine.DEFAULT_REFERENCE_BACKGROUND_COL);
            }else{
                getContentPane().setBackground(CZSystemDefine.DEFAULT_BACKGROUND_COL);
            }

            group_name = new JLabel("�O���[�v",JLabel.CENTER);
            group_name.setBounds(20, 20, 100, 24);
            group_name.setLocale(new Locale("ja","JP"));
            group_name.setFont(new java.awt.Font("dialog", 0, 16));
            group_name.setBorder(new Flush3DBorder());
            group_name.setForeground(java.awt.Color.black);
            getContentPane().add(group_name);

            recip_name = new JLabel("���V�s",JLabel.CENTER);
            recip_name.setBounds(120, 20, 540, 24);
            recip_name.setLocale(new Locale("ja","JP"));
            recip_name.setFont(new java.awt.Font("dialog", 0, 16));
            recip_name.setBorder(new Flush3DBorder());
            recip_name.setForeground(java.awt.Color.black);
            getContentPane().add(recip_name);

            large_no = new JLabel("L_No",JLabel.CENTER);
            large_no.setBounds(20, 44, 100, 24);
            large_no.setLocale(new Locale("ja","JP"));
            large_no.setFont(new java.awt.Font("dialog", 0, 16));
            large_no.setBorder(new Flush3DBorder());
            large_no.setForeground(java.awt.Color.black);
            getContentPane().add(large_no);

            large_name = new JLabel("�區��",JLabel.CENTER);
            large_name.setBounds(120, 44, 540, 24);
            large_name.setLocale(new Locale("ja","JP"));
            large_name.setFont(new java.awt.Font("dialog", 0, 16));
            large_name.setBorder(new Flush3DBorder());
            large_name.setForeground(java.awt.Color.black);
            getContentPane().add(large_name);

            midle_no = new JLabel("������",JLabel.CENTER);
            midle_no.setBounds(20, 70, 100, 24);
            midle_no.setLocale(new Locale("ja","JP"));
            midle_no.setFont(new java.awt.Font("dialog", 0, 16));
            midle_no.setBorder(new Flush3DBorder());
            midle_no.setForeground(java.awt.Color.black);
            getContentPane().add(midle_no);

            midName = new ItemText();
            midName.setBounds(120, 70, 540, 24);
            midName.setLocale(new Locale("ja","JP"));
            midName.setFont(new java.awt.Font("dialog", 0, 16));
            midName.setBorder(new Flush3DBorder());
            midName.setForeground(java.awt.Color.black);
            getContentPane().add(midName);

            // �I�y���[�^��
            JLabel lab = null;
            lab = new JLabel("�ݒ��",JLabel.CENTER);
            lab.setBounds(20, 132, 100, 24);
            lab.setLocale(new Locale("ja","JP"));
            lab.setFont(new java.awt.Font("dialog", 0, 16));
            lab.setBorder(new Flush3DBorder());
            lab.setForeground(java.awt.Color.black);
            getContentPane().add(lab);

            op_name = new TText();
            op_name.setBounds(120, 132, 140, 24);
            getContentPane().add(op_name);

            midSetWinSendButton = new JButton("��  �s");
            midSetWinSendButton.setBounds(260, 132, 100, 24);
            midSetWinSendButton.setLocale(new Locale("ja","JP"));
            midSetWinSendButton.setFont(new java.awt.Font("dialog", 0, 18));
            midSetWinSendButton.setBorder(new Flush3DBorder());
            midSetWinSendButton.setForeground(java.awt.Color.black);
            midSetWinSendButton.addActionListener(new MidSetWinSendButton());
            getContentPane().add(midSetWinSendButton);

            midSetWinCancelButton = new JButton("�I  ��");
            midSetWinCancelButton.setBounds(560, 132, 100, 24);
            midSetWinCancelButton.setLocale(new Locale("ja","JP"));
            midSetWinCancelButton.setFont(new java.awt.Font("dialog", 0, 18));
            midSetWinCancelButton.setBorder(new Flush3DBorder());
            midSetWinCancelButton.setForeground(java.awt.Color.black);
            midSetWinCancelButton.addActionListener(new MidSetWinCancelButton());
            getContentPane().add(midSetWinCancelButton);
        }

        //
        //
        //
        public boolean setDefault(int grp,String _grp,int rcp,String _rcp,
                                    int lag,String _lag,int mid,String _mid){
            current_group = grp;
            current_recip = rcp;
            current_large = lag;
            current_midle = mid;

//@@            CZSystem.log("CZControlTable","T6MidSetWin [" + _mid + "]");
            group_name.setText(_grp);
            recip_name.setText(_rcp);
            large_no.setText(new String("" + lag + ""));
            large_name.setText(_lag);
            midle_no.setText(new String("" + mid + ""));
            midName.setText(_mid);
            op_name.setText("");
            return true;
        }

        //
        //
        //
        public boolean setT6MidSendStatus(){

            if(T6 != current_group) return false;
            if(1  > current_recip) return false;
            if(999 < current_recip) return false;

            sendOp = op_name.getText();
            if(1 > sendOp.length()){
                return false;
            }

            sendMidName = midName.getText();
            if(1 > sendMidName.length()){
                return false;
            }
            return true;
        }

        /*
        *       ����e�[�u���F�^�C�g������͂���TextField
        */
        public class ItemText extends JTextField {

            ItemText(){
                    super();
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

                //
                //
                public void insertString( int offset, String str, AttributeSet a )
                                                        throws BadLocationException {

                    String tmp = new String(getText(0,getLength()) + str);
                    byte   b[];

                    try{
                        b = tmp.getBytes("SJIS");
                    }
                    catch(Exception e){
                        CZSystem.log("CZControlTable","T6MidSetWin ItemText [" + e + "]");
                        return;
                    }

//@@                    CZSystem.log("CZControlTable","T6MidSetWin ItemText [" + tmp + "][" + b + "][" + b.length + "]");

                    if(64 < b.length) return;
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

        /**
        *       �ݒ�҂���͂���TextField
        */
        public class TText extends JTextField {

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

        /*
        *
        *
        *
        */
        class MidSetWinSendButton implements ActionListener {

            public void actionPerformed(ActionEvent ev){

                if(!setT6MidSendStatus()){
                    Object msg[] = {"����e�[�u��:�s�U�����ڕύX",
                                    "�O���[�v�A���V�s�[�A�ݒ�ҁA�^�C�g����",
                                    "�������Ă�������"};
                    errorMsg(msg);
                    return;
                }

                //Send
//@@                CZSystem.log("CZControlTable","T6LimitWin MidSetWinSendButton-->[" +
//@@                        sendOp + "][" + current_group + "][" + current_recip + "][" + sendMidName + "]");

/*@@
                if(!CZSystem.CZControlT6MidExchange(sendOp,current_group,
                            current_recip,current_large,current_midle,sendMidName)){

                    Object msg[] = {"����e�[�u��:�s�U�����ڕύX",
                                    "�ύX�����s���܂���",
                                    "�Ǘ��҂ɖ₢���킹�Ă�������"};
                    errorMsg(msg);
                    return;
                }
@@*/
                return ;
            }
        }

        /*
        *
        *
        *
        */
        class MidSetWinCancelButton implements ActionListener {
            public void actionPerformed(ActionEvent ev){
                setVisible(false);
            }
        }
    }  //public class T6MidSetWin extends JDialog

    /*
    *
    *       ���ڃ����W�ύX�pWindow
    *
    */
    public class T6LimitWin extends JDialog {

        private CZSystemCtT6Name  ctT6Name = null;

//        private ItemText    item_name   = null;
        private JComboBox   sort_kubun  = null;
        private PVText      pv_no       = null;

        private ItemText    k_name      = null;
        private MinMaxText  k_min       = null;
        private MinMaxText  k_max       = null;
        private DigitText   k_digit     = null;
        private UnitText    k_unit      = null;

        private TText       op_name     = null;

        private JButton     unit_send_button   = null;
        private JButton     unit_cancel_button = null;


        private String      sendOp    = null;
        private String      sendName  = null;
        private float       sendMin   = 0.0f;
        private float       sendMax   = 1.0f;
        private String      sendUnit  = null;
        private int         sendDigit = 0;
        private int         sendPVNo  = 1;
        private int         sendSort  = 1;

        //
        //
        //
        T6LimitWin(){
            super();

            setTitle("����e�[�u��:�s�U���ڐݒ�");
            setSize(715,160);
            setResizable(false);
            setModal(true);

            getContentPane().setLayout(null);
            // ����n�Q�Ƌ@�\    @20131021
            if(CZSystemDefine.REFERENCE_RUN == CZSystem.getRunLevel()){
                getContentPane().setBackground(CZSystemDefine.DEFAULT_REFERENCE_BACKGROUND_COL);
            }else{
                getContentPane().setBackground(CZSystemDefine.DEFAULT_BACKGROUND_COL);
            }

            JLabel lab = null;

            lab = new JLabel("��                ��",JLabel.CENTER);
            lab.setBounds(20, 20, 300, 24);
            lab.setLocale(new Locale("ja","JP"));
            lab.setFont(new java.awt.Font("dialog", 0, 16));
            lab.setBorder(new Flush3DBorder());
            lab.setForeground(java.awt.Color.black);
            getContentPane().add(lab);

            lab = new JLabel("�l����",JLabel.CENTER);
            lab.setBounds(320, 20, 100, 24);
            lab.setLocale(new Locale("ja","JP"));
            lab.setFont(new java.awt.Font("dialog", 0, 16));
            lab.setBorder(new Flush3DBorder());
            lab.setForeground(java.awt.Color.black);
            getContentPane().add(lab);

            lab = new JLabel("�l����",JLabel.CENTER);
            lab.setBounds(420, 20, 100, 24);
            lab.setLocale(new Locale("ja","JP"));
            lab.setFont(new java.awt.Font("dialog", 0, 16));
            lab.setBorder(new Flush3DBorder());
            lab.setForeground(java.awt.Color.black);
            getContentPane().add(lab);

            lab = new JLabel("��",JLabel.CENTER);
            lab.setBounds(520, 20, 25, 24);
            lab.setLocale(new Locale("ja","JP"));
            lab.setFont(new java.awt.Font("dialog", 0, 16));
            lab.setBorder(new Flush3DBorder());
            lab.setForeground(java.awt.Color.black);
            getContentPane().add(lab);

            lab = new JLabel("�P        ��",JLabel.CENTER);
            lab.setBounds(545, 20, 150, 24);
            lab.setLocale(new Locale("ja","JP"));
            lab.setFont(new java.awt.Font("dialog", 0, 16));
            lab.setBorder(new Flush3DBorder());
            lab.setForeground(java.awt.Color.black);
            getContentPane().add(lab);

            k_name = new ItemText();
            k_name.setBounds(20, 44, 300, 24);
            k_name.setLocale(new Locale("ja","JP"));
            k_name.setFont(new java.awt.Font("dialog", 0, 16));
            k_name.setBorder(new Flush3DBorder());
            k_name.setForeground(java.awt.Color.black);
            getContentPane().add(k_name);

            k_min = new MinMaxText();
            k_min.setBounds(320, 44, 100, 24);
            k_min.setLocale(new Locale("ja","JP"));
            k_min.setFont(new java.awt.Font("dialog", 0, 16));
            k_min.setBorder(new Flush3DBorder());
            k_min.setForeground(java.awt.Color.black);
            getContentPane().add(k_min);

            k_max = new MinMaxText();
            k_max.setBounds(420, 44, 100, 24);
            k_max.setLocale(new Locale("ja","JP"));
            k_max.setFont(new java.awt.Font("dialog", 0, 16));
            k_max.setBorder(new Flush3DBorder());
            k_max.setForeground(java.awt.Color.black);
            getContentPane().add(k_max);

            k_digit = new DigitText();
            k_digit.setBounds(520, 44, 25, 24);
            k_digit.setLocale(new Locale("ja","JP"));
            k_digit.setFont(new java.awt.Font("dialog", 0, 16));
            k_digit.setBorder(new Flush3DBorder());
            k_digit.setForeground(java.awt.Color.black);
            getContentPane().add(k_digit);

            k_unit = new UnitText();
            k_unit.setBounds(545, 44, 150, 24);
            k_unit.setLocale(new Locale("ja","JP"));
            k_unit.setFont(new java.awt.Font("dialog", 0, 16));
            k_unit.setBorder(new Flush3DBorder());
            k_unit.setForeground(java.awt.Color.black);
            getContentPane().add(k_unit);

            // �I�y���[�^��
            lab = new JLabel("�ݒ��",JLabel.CENTER);
            lab.setBounds(20, 86, 100, 24);
            lab.setLocale(new Locale("ja","JP"));
            lab.setFont(new java.awt.Font("dialog", 0, 16));
            lab.setBorder(new Flush3DBorder());
            lab.setForeground(java.awt.Color.black);
            getContentPane().add(lab);
            op_name = new TText();
            op_name.setBounds(120, 86, 140, 24);
            getContentPane().add(op_name);

            unit_send_button = new JButton("��  �s");
            unit_send_button.setBounds(260, 86, 100, 24);
            unit_send_button.setLocale(new Locale("ja","JP"));
            unit_send_button.setFont(new java.awt.Font("dialog", 0, 18));
            unit_send_button.setBorder(new Flush3DBorder());
            unit_send_button.setForeground(java.awt.Color.black);
            unit_send_button.addActionListener(new UnitSendButton());
            getContentPane().add(unit_send_button);

            unit_cancel_button = new JButton("�I  ��");
            unit_cancel_button.setBounds(595, 86, 100, 24);
            unit_cancel_button.setLocale(new Locale("ja","JP"));
            unit_cancel_button.setFont(new java.awt.Font("dialog", 0, 18));
            unit_cancel_button.setBorder(new Flush3DBorder());
            unit_cancel_button.setForeground(java.awt.Color.black);
            unit_cancel_button.addActionListener(new UnitCancelButton());
            getContentPane().add(unit_cancel_button);

        }

        //
        //
        //
        public boolean setDefault(CZSystemCtT6Name _name){

//@@            CZSystem.log("CZControlTable","T6LimitWin setDefault()");

            if(null == _name) return false;
            ctT6Name = _name;

            k_name.setText(ctT6Name.k_name.trim());
            k_min.setText(Float.toString(ctT6Name.k_min));
            k_max.setText(Float.toString(ctT6Name.k_max));
            k_digit.setText(Integer.toString(ctT6Name.k_keta));
            k_unit.setText(ctT6Name.k_unit.trim());

            op_name.setText("");

            return true;
        }

        //
        //
        //
        private boolean setUnitSendStatus(){
            sendOp = op_name.getText();
            if(1 > sendOp.length()){
                return false;
            }

            sendName = k_name.getText();
            if(1 > sendName.length()){
                return false;
            }

            sendUnit = k_unit.getText();
            if(1 > sendUnit.length()){
                return false;
            }

            try{
                sendMin   = Float.parseFloat(k_min.getText());
                sendMax   = Float.parseFloat(k_max.getText());
                sendDigit = Integer.parseInt(k_digit.getText());
            }
            catch (Exception e){
                return false;
            }

            if(sendMin >= sendMax) return false;
            if(0 > sendDigit) return false;

            return true;
        }

        /*
        *       ����e�[�u���F���ږ�����͂���TextField
        */
        public class ItemText extends JTextField {

            ItemText(){
                super();
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

                //
                //
                public void insertString( int offset, String str, AttributeSet a )
                                                throws BadLocationException {

                    String tmp = new String(getText(0,getLength()) + str);
                    byte   b[];

                    try{
                        b = tmp.getBytes("SJIS");
                    }
                    catch(Exception e){
                        CZSystem.log("CZControlTable","T6LimitWin ItemText [" + e + "]");
                        return;
                    }

//@@                    CZSystem.log("CZControlTable","T6LimitWin ItemText [" + tmp + "][" + b + "][" + b.length + "]");

//@@@                    if(32 < b.length) return;
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


        /*
        *       ����e�[�u���F�Ή��o�u����͂���TextField
        */
        public class PVText extends JTextField {

            PVText(){
                super();
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

                    if(2 < getLength()) return;
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

        /**
        *       ����e�[�u���F�l�����l��������͂���TextField
        */
        public class MinMaxText extends JTextField {

            MinMaxText(){
                super();
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
                String validValues = "0123456789.-";

                //
                //
                public void insertString( int offset, String str, AttributeSet a )
                                    throws BadLocationException {

                    if(9 < getLength()) return;
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

        /**
        *       ����e�[�u���F������͂���TextField
        */
        public class DigitText extends JTextField {

            DigitText(){
                super();
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
                String validValues = "0123456";

                //
                //
                public void insertString( int offset, String str, AttributeSet a )
                                        throws BadLocationException {

                    if(0 < getLength()) return;
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

        /*
        *       ����e�[�u���F�P�ʂ���͂���TextField
        */
        public class UnitText extends JTextField {

            UnitText(){
                super();
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

                //
                //
                public void insertString( int offset, String str, AttributeSet a )
                                                throws BadLocationException {

                    String tmp = new String(getText(0,getLength()) + str);
                    byte   b[];

                    try{
                        b = tmp.getBytes("SJIS");
                    }
                    catch(Exception e){
                        CZSystem.log("CZControlTable","T6LimitWin ItemText Error [" + e + "]");
                        return;
                    }

//@@                    CZSystem.log("CZControlTable","T6LimitWin ItemText [" + tmp + "][" + b + "][" + b.length + "]");

                    if(16 < b.length) return;
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

        /*
        *
        *       �ݒ�҂���͂���TextField
        *
        */
        public class TText extends JTextField {

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

        /*
        *
        *
        *
        */
        class UnitSendButton implements ActionListener {
            public void actionPerformed(ActionEvent ev){

                if(!setUnitSendStatus()){
                    Object msg[] = {"����e�[�u��:�s�U���ڍX�V",
                                    "�ݒ�ҁA���ځAMin�AMax�A����",
                                    "�������Ă�������"};
                    errorMsg(msg);
                    return;
                }

//@@                CZSystem.log("CZControlTable","UnitSendButton-->[" + ctT6Name.g_no  + "][" + ctT6Name.k_no1 + "]["
//@@                                                         ctT6Name.k_no2   + "][" +
//@@                                                         sendOp    + "][" + sendName  + "][" +
//@@                                                         sendUnit  + "][" +
//@@                                                         sendMin   + "][" + sendMax  + "][" +
//@@                                                         sendDigit + "]");

                CZParamControlT6Define s = new CZParamControlT6Define();
                s.setName(sendName);
                s.setTani(sendUnit);
                s.setMin(sendMin);
                s.setMax(sendMax);
                s.setPoint(sendDigit);
                //Send
                if(!CZSystem.CZControlT6DefineExchange(sendOp, ctT6Name.g_no , ctT6Name.k_no1, ctT6Name.k_no2, ctT6Name.k_no, s)){

                    Object msg[] = {"����e�[�u��:�s�U���ڍX�V",
                                "�X�V�����s���܂���",
                                "�Ǘ��҂ɖ₢���킹�Ă�������"};
                    errorMsg(msg);
                    return;
                }

                return ;
            }
        }


        /*
        *
        *
        *
        */
        class UnitCancelButton implements ActionListener {
            public void actionPerformed(ActionEvent ev){
                setVisible(false);
            }
        }
    } //public class T6LimitWin extends JDialog

    /**
    *
    *       ���ڐݒ�pWindow
    *
    */
    public class T6SetWin extends JDialog {

        private TText       op_name      = null;

        private JButton     updateButton = null;
        private JButton     saveButton   = null;
        private JButton     cancelButton = null;

        /****** 2007.01.24 ADD ******/
        private JLabel      rcp_no_lab   = null;

        private JLabel      status_name  = null;

        private JButton     statusButton = null;
        /****************************/

        private T6Table     t6Table      = null;     //T6����

        private String      sendOp    = null;
        private float       sendVal   = 0.0f;

        private int       rcp   = 0;
        private int       lag   = 0;
        private int       mid   = 0;
        private int       dataCount  = 0;


        //
        //
        //
        T6SetWin(){
            super();

            setTitle("����e�[�u��:�s�U���ڐݒ�");
    	    /****** 2007.01.24 ADD ******/
            setSize(735,530);
/*			setSize(805,530);*/
	        /****************************/
            setResizable(false);
            setModal(true);

            addWindowListener(new WindowAdapter(){
                public void windowClosing(WindowEvent e){
                //    winClose(e);
                }
            });

            getContentPane().setLayout(null);
            // ����n�Q�Ƌ@�\    @20131021
            if(CZSystemDefine.REFERENCE_RUN == CZSystem.getRunLevel()){
                getContentPane().setBackground(CZSystemDefine.DEFAULT_REFERENCE_BACKGROUND_COL);
            }else{
                getContentPane().setBackground(CZSystemDefine.DEFAULT_BACKGROUND_COL);
            }

            //T6���ڃe�[�u�� @@
            t6Table = new T6Table();
            JTableHeader tabHead = t6Table.getTableHeader();
            tabHead.setReorderingAllowed(false);
            JScrollPane panel = new JScrollPane(t6Table);
    	    /****** 2007.01.24 ADD ******/
            panel.setBounds(20, 50, 670, 380);
            getContentPane().add(panel);

            JLabel lab = null;
            // �I�y���[�^��
            lab = new JLabel("�ݒ��",JLabel.CENTER);
    	    /****** 2007.01.24 ADD ******/
            lab.setBounds(20, 460, 100, 24);
            lab.setLocale(new Locale("ja","JP"));
            lab.setFont(new java.awt.Font("dialog", 0, 16));
            lab.setBorder(new Flush3DBorder());
            lab.setForeground(java.awt.Color.black);
            getContentPane().add(lab);

            op_name = new TText();
    	    /****** 2007.01.24 ADD ******/
            op_name.setBounds(120, 460, 140, 24);
            getContentPane().add(op_name);

            updateButton = new JButton();
            updateButton = new JButton("�C  ��");
    	    /****** 2007.01.24 ADD ******/
            updateButton.setBounds(340, 460, 100, 24);
            updateButton.setLocale(new Locale("ja","JP"));
            updateButton.setFont(new java.awt.Font("dialog", 0, 18));
            updateButton.setBorder(new Flush3DBorder());
            updateButton.setForeground(java.awt.Color.black);
            updateButton.addActionListener(new UpdateSendButton());
            getContentPane().add(updateButton);

            saveButton = new JButton();
//            saveButton = new JButton("�C���ۑ�");
            saveButton = new JButton("��  ��");
    	    /****** 2007.01.24 ADD ******/
            saveButton.setBounds(460, 460, 100, 24);
            saveButton.setLocale(new Locale("ja","JP"));
            saveButton.setFont(new java.awt.Font("dialog", 0, 18));
            saveButton.setBorder(new Flush3DBorder());
            saveButton.setForeground(java.awt.Color.black);
            saveButton.addActionListener(new SaveSendButton());
            getContentPane().add(saveButton);

            cancelButton = new JButton("�I  ��");
    	    /****** 2007.01.24 ADD ******/
            cancelButton.setBounds(585, 460, 100, 24);
            cancelButton.setLocale(new Locale("ja","JP"));
            cancelButton.setFont(new java.awt.Font("dialog", 0, 18));
            cancelButton.setBorder(new Flush3DBorder());
            cancelButton.setForeground(java.awt.Color.black);
            cancelButton.addActionListener(new T6SetCancelButton());
            getContentPane().add(cancelButton);

            /****** 2007.01.24 ADD *******/
            rcp_no_lab = new JLabel("",JLabel.CENTER);
            rcp_no_lab.setBounds(20, 20, 80, 24);
            rcp_no_lab.setLocale(new Locale("ja","JP"));
            rcp_no_lab.setFont(new java.awt.Font("dialog", 0, 18));
            rcp_no_lab.setBorder(new Flush3DBorder());
            rcp_no_lab.setForeground(java.awt.Color.black);
            getContentPane().add(rcp_no_lab);

            status_name = new JLabel("�J�����g�\��",JLabel.CENTER);
            status_name.setBounds(120, 20, 200, 24);
            status_name.setLocale(new Locale("ja","JP"));
            status_name.setFont(new java.awt.Font("dialog", 0, 18));
            status_name.setBorder(new Flush3DBorder());
            status_name.setForeground(java.awt.Color.black);
            getContentPane().add(status_name);

            statusButton = new JButton("�\���ؑ�");
            statusButton.setBounds(320, 20, 150, 24);
            statusButton.setLocale(new Locale("ja","JP"));
            statusButton.setFont(new java.awt.Font("dialog", 0, 18));
            statusButton.setBorder(new Flush3DBorder());
            statusButton.setForeground(java.awt.Color.black);
            statusButton.addActionListener(new T6SetStatusButton());
            getContentPane().add(statusButton);
            /*****************************/
        }

        //
        //
        //
        public boolean setDefault(int g, int r, int m, int l){

            CZSystem.log("CZControlTable","T6SetWin setDefault()");

            t6Table.setData( g, r, m,  l);
            op_name.setText("");

            
            rcp_no_lab.setText(Integer.toString(r));

            /****** 2007.01.24 ADD ******/
            if(Status_flg == false){
                status_name.setText("�}�X�^�[�\��");
            } else {
				status_name.setText("�J�����g�\��");
			}

			if(haita_flg == false){
				updateButton.setEnabled(false);
				saveButton.setEnabled(false);
				/* 2007.03.13 y.k start */
				if(Button_flg == false){
					statusButton.setEnabled(false);
				} else {
					statusButton.setEnabled(true);
				}
				/* 2007.03.13 y.k end */
			} else {
				if(Button_flg == false){
					saveButton.setEnabled(true);	/* 2007.03.13 y.k */
					updateButton.setEnabled(false);
					statusButton.setEnabled(false);
				} else {
					saveButton.setEnabled(true);	/* 2007.03.13 y.k */
					updateButton.setEnabled(true);
					statusButton.setEnabled(true);
				}
			}
            /****************************/

            return true;
        }

        //
        //
        //
        private boolean setSendStatus(){

/* *** 2007.02.02 * y.k *** */
            if(t6Table.isEditing()){
                System.out.println("EditColumn :[" + t6Table.getEditingColumn() + "] ROW:["+ t6Table.getEditingRow() + "]");
                CZSystem.log("CZControlTable UpdateSendButton or SaveSendButton"," actionPerformed Table Data EDIT !!");
                Object msg[] = {"����e�[�u��",
                                "�ݒ蒆���ڗL��I�I",
                                ""};
                errorMsg(msg);
                return false;
            }
/* *** 2007.02.02 * y.k *** */

            // �ݒ�҃`�F�b�N
            sendOp = op_name.getText();
            if(1 > sendOp.length()){
                CZSystem.log("CZControlTable UpdateSendButton or SaveSendButton",
                             "actionPerformed Table Op Name Error !!");
                Object msg[] = {"����e�[�u��",
                                "�ݒ�҂���͂��Ă��������I�I",
                                ""};
                errorMsg(msg);
                return false;
            }
            // Min,Max�`�F�b�N
            float min = 0.0f;
            float max = 0.0f;
            float val = 0.0f;
            for (int i=0; i<dataCount; i++) {
                min = ((Float) t6Table.getValueAt(i, 2)).floatValue();
                max = ((Float) t6Table.getValueAt(i, 3)).floatValue();
                val = ((Float) t6Table.getValueAt(i, 7)).floatValue();
//@@                System.out.println("Min "+ min + ":Max "+ max + ":Val "+ val);

                if ((min > val) || (max < val)) {
                    Object msg[] = {"����e�[�u��",
                                    "�ݒ�l���l�����A�l�����͈̔͊O�ł��I�I",
                                    ""};
                    errorMsg(msg);
                    return false;
                }

            }
            return true;
        }

        /**
         *
         *      �I��
         *
         */
        class T6SetCancelButton implements ActionListener {
            public void actionPerformed(ActionEvent ev){
                setT6Win_.setVisible(false);
            }
        } //class T6SetCancelButton implements ActionListener

        /**
         *
         *      �C��
         *
         */
        class UpdateSendButton implements ActionListener {
            public void actionPerformed(ActionEvent ev){
                if (setSendStatus()) {

                    CZParamT6Table[] param = new CZParamT6Table[dataCount];
                    for (int i=0; i < dataCount; i++) {
                        param[i] = new CZParamT6Table();
                        param[i].setGrpNo(6);
                        param[i].setRcpNo(rcp);
                        param[i].setLagNo(lag);
                        param[i].setMidNo(mid);
                        param[i].setKNo(((Integer)t6Table.getValueAt(i, 0)).intValue());
                        param[i].setVal(((Float) t6Table.getValueAt(i, 7)).floatValue());
//@@                        System.out.println("" + ((Float) t6Table.getValueAt(i, 7)).floatValue());
                    }
                    CZSystem.CZControlT6TableExchange(1, op_name.getText(), param);
                }
            }
        } //class UpdateSendButton implements ActionListener

        /**
         *
         *      �ۑ�
         *
         */
        class SaveSendButton implements ActionListener {
            public void actionPerformed(ActionEvent ev){
                if (setSendStatus()) {

                    CZParamT6Table[] param = new CZParamT6Table[dataCount];
                    for (int i=0; i < dataCount; i++) {
                        param[i] = new CZParamT6Table();
                        param[i].setGrpNo(6);
                        param[i].setRcpNo(rcp);
                        param[i].setLagNo(lag);
                        param[i].setMidNo(mid);
                        param[i].setKNo(((Integer)t6Table.getValueAt(i, 0)).intValue());
                        param[i].setVal(((Float) t6Table.getValueAt(i, 7)).floatValue());
//@@                        System.out.println("" + ((Float) t6Table.getValueAt(i, 7)).floatValue());
                    }
/* 2004.08.06 �C���E�ۑ��@�Ɨ��Ή�       CZSystem.CZControlT6TableExchange(2, op_name.getText(), param); */
                    CZSystem.CZControlT6TableExchange(0, op_name.getText(), param);
                }
            }
        } //class SaveSendButton implements ActionListener

        /**
         *
         *      �\���؂�ւ� 2007.01.24 ADD
         *
         */
        class T6SetStatusButton implements ActionListener {
            public void actionPerformed(ActionEvent ev){
                
                if(Status_flg == true){
					Status_flg = false;
				} else {
					Status_flg = true;
				}

        if( 0 != CZSystemDefine.TIMER_FLG ){
			at.stop();
			t.restart();
			CZSystem.log("CZControlTable","Restart Alarm Display Open(T6 Setting Change)");
		}

            int gNo = c_table.getSelectedRow();
            int rNo = g_table.getSelectedRow();
            int lNo = t6LagTable_.getSelectedRow();
            int mNo = t6MidTable_.getSelectedRow();

            if(0 > gNo) return;
            if(0 > rNo) return;
            if(0 > lNo) return;
            if(0 > mNo) return;
            gNo++;
            rNo++;
            lNo++;
            mNo++;

            Integer group  = (Integer)c_table.getValueAt(c_table.getSelectedRow(),1);
            Integer recip  = (Integer)g_table.getValueAt(g_table.getSelectedRow(),0);

            boolean current = Status_flg;

            t6Current_ = CZSystem.getCtT6Tb(6, rNo, lNo, mNo, current);

            CZSystem.log("CZControlTable","Current or Master [" + current + "]");

            setT6Win_.setDefault( gNo, rNo, lNo, mNo );

            setT6Win_.setVisible(true);

            return ;
			}
		}

        /**
        *
        *   ����T6���ڃe�[�u���N���X
        *
        */
        public class T6Table extends JTable {

            private T6TblMdl model = null;

            T6Table(){
                super();

                try{
                    setName("T6TblMdl");
                    setBounds(0, 0, 200, 200);
                    setAutoCreateColumnsFromModel(true);
                    setSelectionMode(ListSelectionModel.SINGLE_SELECTION);

                    setLocale(new Locale("ja","JP"));
                    setFont(new java.awt.Font("dialog", 0, 12));
                    setRowHeight(17);

                    model = new T6TblMdl();

                    setModel(model);


                    DefaultTableColumnModel cmdl = (DefaultTableColumnModel)getColumnModel();
                    TableColumn  colum = null;
                    T6TblRenderer ren  = null;

                    //#
                    colum = cmdl.getColumn(0);
                    colum.setMaxWidth(40);
                    colum.setMinWidth(40);
                    colum.setWidth(40);
                    ren = new T6TblRenderer();
                    ren.setHorizontalAlignment(ren.RIGHT);
                    colum.setCellRenderer(ren);

                    //���ږ�
                    colum = cmdl.getColumn(1);
                    colum.setMaxWidth(220);
                    colum.setMinWidth(220);
                    colum.setWidth(220);
                    ren = new T6TblRenderer();
                    colum.setCellRenderer(ren);

                    //Min
                    colum = cmdl.getColumn(2);
                    colum.setMaxWidth(60);
                    colum.setMinWidth(60);
                    colum.setWidth(60);
                    ren = new T6TblRenderer();
                    ren.setHorizontalAlignment(ren.RIGHT);
                    colum.setCellRenderer(ren);

                    //Max
                    colum = cmdl.getColumn(3);
                    colum.setMaxWidth(60);
                    colum.setMinWidth(60);
                    colum.setWidth(60);
                    ren = new T6TblRenderer();
                    ren.setHorizontalAlignment(ren.RIGHT);
                    colum.setCellRenderer(ren);

                    //��
                    colum = cmdl.getColumn(4);
                    colum.setMaxWidth(30);
                    colum.setMinWidth(30);
                    colum.setWidth(30);
                    ren = new T6TblRenderer();
                    ren.setHorizontalAlignment(ren.RIGHT);
                    colum.setCellRenderer(ren);

                    //�P��
                    colum = cmdl.getColumn(5);
                    colum.setMaxWidth(100);
                    colum.setMinWidth(100);
                    colum.setWidth(100);
                    ren = new T6TblRenderer();
                    colum.setCellRenderer(ren);

                    //���ݒl
                    colum = cmdl.getColumn(6);
                    colum.setMaxWidth(70);
                    colum.setMinWidth(70);
                    colum.setWidth(70);
                    ren = new T6TblRenderer();
                    ren.setHorizontalAlignment(ren.RIGHT);
                    colum.setCellRenderer(ren);

                    //�ύX�l
                    colum = cmdl.getColumn(7);
                    colum.setMaxWidth(70);
                    colum.setMinWidth(70);
                    colum.setWidth(70);
                    ren = new T6TblRenderer();
                    ren.setHorizontalAlignment(ren.RIGHT);
                    colum.setCellRenderer(ren);
                }
                catch (Throwable e) {
                    CZSystem.handleException(e);
                }
            }

            public void setData(int gr, int rp, int lg, int md){

                rcp = rp;
                lag = lg;
                mid = md;
//@@                CZSystem.log("CZControlTable","T6Table setData [" + gr + "][" + lg + "][" + md + "]");

                CZSystemCtT6Name name =  null;
                CZSystemCtT6Tb   data =  null;
                String      empty   = "";

                dataCount = 0;
                for(int i = 0 ; i < 600 ; i++){
                    name = CZSystem.getCtT6Name(gr, lg, md, i+1);
                    data = getCtT6Tb(gr, rp, lg, md, i+1);

                    if(null != name){
                        setValueAt(name.k_name.trim(),i,1);
                        setValueAt(new Float(name.k_min),i,2);
                        setValueAt(new Float(name.k_max),i,3);
                        setValueAt(new Integer(name.k_keta),i,4);
                        setValueAt(name.k_unit.trim(),i,5);
                        if ( null != data) {
                            setValueAt(new Float(data.k_val),i,6);
                            setValueAt(new Float(data.k_val),i,7);
                        } else {
                            setValueAt(new Float(0.0f),i,6);
                            setValueAt(new Float(0.0f),i,7);
                        }
                        dataCount++;
                    }
                    else {
                        setValueAt(empty,i,1);
                        setValueAt(empty,i,2);
                        setValueAt(empty,i,3);
                        setValueAt(empty,i,4);
                        setValueAt(empty,i,5);
                        setValueAt(empty,i,6);
                        setValueAt(empty,i,7);

                    }
                } // for end
                setRowSelectionInterval(0,0);

                Rectangle cellRect = getCellRect(0,0,false);
                if(cellRect != null){
                    scrollRectToVisible(cellRect);
                }
                repaint();
            }
        }

        /*
        *
        *   ����T6���ڃe�[�u���N���X�F���f��
        *
        */

        public class  T6TblMdl extends AbstractTableModel {

            final   int TBL_COL = 8;
            private int TBL_ROW = 600;

            private Object data[][];

            final String[] names = {"#", "��    ��",
                        "Min","Max","��","�P��","���ݒl","�ύX�l"};

            T6TblMdl(){
                super();

                data = new Object[TBL_ROW][TBL_COL];

                try{

                    for(int i = 0 ; i < TBL_ROW ; i++){
                        data[i][0] = new Integer(i+1);
                        data[i][1] = new String("################################");
                        data[i][2] = new String("#.#####");
                        data[i][3] = new String("#####.#####");
                        data[i][5] = new String("######");
                        data[i][6] = new String("######");
                        data[i][7] = new String("######");
                    }
                }
                catch (Throwable e) {
                    CZSystem.handleException(e);
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
                if(7 == col) return true;
                return false;
            }

            public void setValueAt(Object aValue, int row, int column){
                data[row][column] = aValue;
            }
        }


        /*
        *
        *   ����T6���ڃe�[�u���N���X�F�����_���[
        *
        *
        */

        class T6TblRenderer extends DefaultTableCellRenderer {

            T6TblRenderer(){
            super();
                setLocale(new Locale("ja","JP"));
                setFont(new java.awt.Font("dialog", 0, 12));
            }

            public Component getTableCellRendererComponent( JTable table,
                                                            Object value,
                                                            boolean isSelected,
                                                            boolean hasFocus,
                                                            int row,int column){


                if(6 != column){
                    super.getTableCellRendererComponent(table,
                                                        value,
                                                        isSelected,
                                                        hasFocus,
                                                        row,column);
                    return(this);
                }

                if(Float.class != table.getValueAt(row, 2).getClass()){
                    super.getTableCellRendererComponent(table,
                                                        value,
                                                        isSelected,
                                                        hasFocus,
                                                        row,column);
                    return(this);
                }

                if(Float.class != table.getValueAt(row, 3).getClass()){
                    super.getTableCellRendererComponent(table,
                                                        value,
                                                        isSelected,
                                                        hasFocus,
                                                        row,column);
                    return(this);
                }

                if(Integer.class != table.getValueAt(row, 4).getClass()){
                    super.getTableCellRendererComponent(table,
                                                        value,
                                                        isSelected,
                                                        hasFocus,
                                                        row,column);
                    return(this);
                }


                Float   min  = (Float)table.getValueAt(row,2);
                Float   max  = (Float)table.getValueAt(row,3);
                Integer keta = (Integer)table.getValueAt(row,4);
                Float   val  = (Float)table.getValueAt(row,6);

                if(null == val) {
                    val = new Float(0.0f);
                }

                DecimalFormat format = null;
                StringBuffer  buff = new StringBuffer();

                if(1 > keta.intValue()){
                    format = new DecimalFormat("0");
                }
                else {
                    buff.append("0.");
                    for(int i = 0 ; i < keta.intValue() ; i++){
                        buff.append("0");
                    }
                    format = new DecimalFormat(buff.toString());
                }

                Float new_val = new Float(format.format(val));

                super.getTableCellRendererComponent(table,
                                                    format.format(new_val.floatValue()),
                                                    isSelected,
                                                    hasFocus,
                                                    row,column);

                table.setValueAt(new_val,row,column);

                if((min.floatValue() <= new_val.floatValue()) &&
                    (max.floatValue() >= new_val.floatValue())){
                    setForeground(java.awt.Color.blue);
                }
                else {
                    setForeground(java.awt.Color.red);
                }
                return(this);
            }
        }


        /*
        *
        *       �ݒ�҂���͂���TextField
        *
        */
        public class TText extends JTextField {

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

        //
        //  ����e�[�u��(T6)�F�e�[�u�����o��
        //
        private CZSystemCtT6Tb getCtT6Tb(int grp , int rcp, int lag, int mid, int kNo){

            if (null == t6Current_) return null;
            for(int i = 0 ; i < t6Current_.size() ; i++){
                CZSystemCtT6Tb ret = (CZSystemCtT6Tb)t6Current_.elementAt(i);
                if((ret.g_no  == grp) &&
                   (ret.r_no  == rcp) &&
                   (ret.k_no1 == lag) &&
                   (ret.k_no2 == mid) &&
                   (ret.k_no  == kNo)) return ret;
            }
            return null;
        }
    } //public class T6SetWin extends JDialog
// add start 2008.09.12
    /***************************************************************************
    *
    *       �O���[�v����͂���TextField
    *
    ***************************************************************************/
    class RcpText extends JTextField {

        /**
        *
        */
        RcpText(){
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

                if(2 < getLength()) return;
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
     * ���V�s����
     *
     *******************************************************/

    class RecipeAction implements ActionListener {
        public void actionPerformed(ActionEvent ev){

            Integer row = Integer.valueOf(rcp_no_txt.getText());

            
            if(row == 0) {
            }
            else if(row <= 5) {
                g_table.setRowSelectionInterval(0,row-1);
                this.setVerticalScrollBarPosition(0);
            }
            else if((995<= row) && (row < 1000)) {
                g_table.setRowSelectionInterval(0,row-1);
                this.setVerticalScrollBarPosition(17000);
            }
            else {
                g_table.setRowSelectionInterval(0,row-1);
                this.setVerticalScrollBarPosition((row*17)-102);
            } 
        }

        public void setVerticalScrollBarPosition(int position) {
            JScrollBar rcp_jsb = rcp_pnl.getVerticalScrollBar();
            rcp_jsb.setValue(position);
            rcp_pnl.setVerticalScrollBar(rcp_jsb);
        }

    }
// add end 2008.09.12
// add start 2008.09.16
    /***************************************************************************
    *
    *       ���ڔԍ�����͂���TextField
    *
    ***************************************************************************/
    class KoumokuText extends JTextField {

        /**
        *
        */
        KoumokuText(){
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

                if(2 < getLength()) return;
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
     * ���ڔԍ�����
     *
     *******************************************************/

    class KoumokuAction implements ActionListener {
        public void actionPerformed(ActionEvent ev){

            Integer row = Integer.valueOf(koumoku_no_txt.getText());

            
            if((row == 0) || row>600) {
            }
            else if(row <= 5) {
//2011.04.12 Y.K rep
//                v_table.setRowSelectionInterval(0,row-1);
                v_table.setRowSelectionInterval(row-1,row-1);
                this.setVerticalScrollBarPosition(0);
            }
            else if((595<= row) && (row <= 600)) {
//2011.04.12 Y.K rep
//                v_table.setRowSelectionInterval(0,row-1);
                v_table.setRowSelectionInterval(row-1,row-1);
                this.setVerticalScrollBarPosition(10200);
            }
            else {
//2011.04.12 Y.K rep
//                v_table.setRowSelectionInterval(0,row-1);
                  v_table.setRowSelectionInterval(row-1,row-1);
                this.setVerticalScrollBarPosition((row*17)-102);
            } 
        }

        public void setVerticalScrollBarPosition(int position) {
            JScrollBar kmk_jsb = kmk_pnl.getVerticalScrollBar();
            kmk_jsb.setValue(position);
            kmk_pnl.setVerticalScrollBar(kmk_jsb);
        }

    }
// add end 2008.09.16
}
