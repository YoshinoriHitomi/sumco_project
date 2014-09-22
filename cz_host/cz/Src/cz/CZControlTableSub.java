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
 *   ����e�[�u���ύXWindow
 * @author  (KPK Co.,Ltd.)
 * @version 1.0 (2003/04/01) @@  �đ�ł̃R���o�[�g
 * @version 1.1 (2004/01/20) @@@ �O���t��ʂ̏k���g���ǉ�
 * 2008.09.10 H.Nagamine ڼ�ߔԍ��\���ǉ�
 *
 ***********************************************************/
public class CZControlTableSub extends JFrame {

    private final int INC_WIDTH     = 236;  // ���̑���   @@@
    private final int INC_HEIGHT    = 240;  // �����̑��� @@@
    private final int BASE_WIDTH    = 590;  // ��̕�   @@@
    private final int BASE_HEIGHT   = 600;  // ��̍��� @@@
    private final int MAGNIFICATION = 5;    // �g��̍ő�{�� @@@

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

    private final int MAIN1_H_T     = 14;   // 15   ���C���q�[�^�[�P���x
    private final int MAIN1_H_T_PF  = 66;   // 67   ���C���q�[�^�[�P���x�v���t�@�C��
    private final int DIA           = 24;   // 25   ���a
    private final int DIA_PF        = 23;   // 24   ���a�v���t�@�C��
    private final int SXL_ST        = 17;   // 18   �����グ���x
    private final int SXL_ST_PF     = 75;   // 76   �����グ���x�v���t�@�C��

    private final Color NEW_PRO_COL = java.awt.Color.green;
    private final Color OLD_PRO_COL = java.awt.Color.red;
    private final Color VAL_PRO_COL = java.awt.Color.white;
    private final Color VAL_COL     = java.awt.Color.orange;

//@@@@@@@@@@@@@@@@@@@@@@@@@
    private final Color MST_COL = java.awt.Color.black;
    private final Color CUR_COL = java.awt.Color.blue;

    private CZSystemCtTb    send_data[];

    private int             edit_group;         //�ΏۃO���[�v
    private int             edit_recip;         //���V�s�[No
    private int             edit_number;        //����No
    private CZSystemCtName  edit_name;
    private boolean         edit_current;
    private boolean         edit_haita_flg;

    private boolean         mst_show;           // @@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@

    private Vector          current_data;       //�ݒ蒆�̃f�[�^
    
    private Vector          master_data;        //�}�X�^�[�f�[�^ @@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@



    private Vector      pv_data_shld    = null; //�V�����_�[�̃f�[�^
    private Vector      pv_data_body    = null; //�{�f�B�[�̃f�[�^

    private JButton     save_button     = null; //�ۑ��{�^��
    private JButton     modify_button   = null; //�C���{�^��
    private JButton     cancel_button   = null; //�I���{�^��

    private TText       op_name         = null; //�I�y���[�^�[��

    private CtOldTable  c_old_table     = null; //�ݒ�l��\������e�[�u��
    private CtTable     c_table         = null; //�ύX�l��\������e�[�u��
    private ShiftText   shift_text      = null; //�V�t�g�����鐔�l
    private ShiftText   l_shift_text      = null; //�V�t�g�����鐔�l 20060529
    private BunText     l_bun_text      = null; //�k��������
    private BunText     r_bun_text      = null; //�q��������

    private JPanel      graph_panel     = null; //�O���t�p�l��
    private LPanel      l_panel         = null; //X���ڐ�
    private RPanel      r_panel         = null; //Y���ڐ�
    private MainPanel   main_panel      = null; //�O���t���C���p�l��

    private LPanelView      l_panelView         = null; //X���ڐ� @@@
    private RPanelView      r_panelView         = null; //Y���ڐ� @@@
    private MainPanelView   main_panelView      = null; //�O���t���C���p�l�� @@@

    private JButton     baseButton      = null; //��{�^�� @@@
    private JButton     reductionButton = null; //�k���{�^�� @@@
    private JButton     expansionButton = null; //�g��{�^�� @@@

    private JButton     mstShowButton = null; //�}�X�^�[�\���{�^�� @@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@


    private int         currWidth       = 590;  // ���݂̕�   @@@
    private int         currHeight      = 600;  // ���݂̍��� @@@

    // �O���t�R���|�[�l���g���X�i
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
    // ---------- �R���X�g���N�^ ---------------------------
    //
    CZControlTableSub(){
        super();

        setTitle("����e�[�u���ݒ�");
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
        // ����n�Q�Ƌ@�\    @20131021
        if(CZSystemDefine.REFERENCE_RUN == CZSystem.getRunLevel()){
            getContentPane().setBackground(CZSystemDefine.DEFAULT_REFERENCE_BACKGROUND_COL);
        }else{
            getContentPane().setBackground(CZSystemDefine.DEFAULT_BACKGROUND_COL);
        }

        JLabel label = null;

        label = new JLabel("�ݒ��",JLabel.CENTER);
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

        modify_button = new JButton("�C  ��");
        modify_button.setBounds(260, 790, 100, 24);         //@@@
//        modify_button.setBounds(260, 800, 100, 24);
        modify_button.setLocale(new Locale("ja","JP"));
        modify_button.setFont(new java.awt.Font("dialog", 0, 18));
        modify_button.setBorder(new Flush3DBorder());
        modify_button.setForeground(java.awt.Color.black);
        modify_button.addActionListener(new ModifyButton());
        getContentPane().add(modify_button);

//        save_button = new JButton("�C���ۑ�");
        save_button = new JButton("�ۑ�");				// 2004.05.27
        save_button.setBounds(360, 790, 100, 24);           //@@@
//        save_button.setBounds(360, 800, 100, 24);
        save_button.setLocale(new Locale("ja","JP"));
        save_button.setFont(new java.awt.Font("dialog", 0, 18));
        save_button.setBorder(new Flush3DBorder());
        save_button.setForeground(java.awt.Color.black);
        save_button.addActionListener(new SaveButton());
        getContentPane().add(save_button);

        cancel_button = new JButton("�I  ��");
        cancel_button.setBounds(630, 790, 100, 24);         //@@@
//        cancel_button.setBounds(630, 800, 100, 24);
        cancel_button.setLocale(new Locale("ja","JP"));
        cancel_button.setFont(new java.awt.Font("dialog", 0, 18));
        cancel_button.setBorder(new Flush3DBorder());
        cancel_button.setForeground(java.awt.Color.black);
        cancel_button.addActionListener(new CancelButton());
        getContentPane().add(cancel_button);

        label = new JLabel("����",JLabel.CENTER);
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
        label = new JLabel("�F��",JLabel.CENTER);
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

        label = new JLabel("�O���[�v",JLabel.CENTER);
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

        label = new JLabel("���V�s",JLabel.CENTER);
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
		
		
		view_lab = new JLabel("�J�����g�\��",JLabel.CENTER);
		view_lab.setBounds(595, 5, 140, 30);
		view_lab.setLocale(new Locale("ja","JP"));
		view_lab.setFont(new java.awt.Font("dialog", 0, 18));
		view_lab.setBorder(new Flush3DBorder());
		view_lab.setForeground(java.awt.Color.black);
		getContentPane().add(view_lab);
		
//        mstShowButton = new JButton("�}�X�^�[�\��");
        mstShowButton = new JButton("�\���ؑ�");
        mstShowButton.setBounds(600, 40, 130, 24);
        mstShowButton.setLocale(new Locale("ja","JP"));
        mstShowButton.setFont(new java.awt.Font("dialog", 0, 18));
        mstShowButton.setBorder(new Flush3DBorder());
        mstShowButton.setForeground(java.awt.Color.black);
        mstShowButton.addActionListener(new MasterShowAction());	//@@@@
        getContentPane().add(mstShowButton);


        // �O���t�p
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

//@@@   ��������ǉ�
        reductionButton = new JButton("�k ��");
        reductionButton.setBounds(530, 15, 50, 24);
        reductionButton.setLocale(new Locale("ja","JP"));
        reductionButton.setFont(new java.awt.Font("dialog", 0, 18));
        reductionButton.setBorder(new Flush3DBorder());
        reductionButton.setForeground(java.awt.Color.black);
        reductionButton.addActionListener(new ReductionAction());
        graph_panel.add(reductionButton);

        baseButton = new JButton("�� ��");
        baseButton.setBounds(580, 15, 50, 24);
        baseButton.setLocale(new Locale("ja","JP"));
        baseButton.setFont(new java.awt.Font("dialog", 0, 18));
        baseButton.setBorder(new Flush3DBorder());
        baseButton.setForeground(java.awt.Color.black);
        baseButton.addActionListener(new StanderdAction());
        graph_panel.add(baseButton);

        expansionButton = new JButton("�g ��");
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

        // �R���|�[�l���g���X�i ����
        GraphComponentListener graphListener = new GraphComponentListener();
        r_panelView.addComponentListener( graphListener );
        l_panelView.addComponentListener( graphListener );
        main_panelView.addComponentListener( graphListener );

        baseButton.setEnabled(false);
        reductionButton.setEnabled(false);
        expansionButton.setEnabled(true);

//@@@�@�����܂�
        // �e�[�u���p
        table_panel = new JPanel();
        table_panel.setLayout(null);
        table_panel.setBounds(745, 20, 385 ,804);
        table_panel.setBorder(new Flush3DBorder());
        table_panel.setBackground(java.awt.Color.gray);
        getContentPane().add(table_panel);

        label = new JLabel("��  ��  �l",JLabel.CENTER);
        label.setBounds(10, 10, 178, 20);
        label.setLocale(new Locale("ja","JP"));
        label.setFont(new java.awt.Font("dialog", 0, 14));
        label.setBorder(new Flush3DBorder());
        label.setForeground(java.awt.Color.black);
        table_panel.add(label);

        label = new JLabel("��  �X  �l",JLabel.CENTER);
        label.setBounds(197, 10, 178, 20);
        label.setLocale(new Locale("ja","JP"));
        label.setFont(new java.awt.Font("dialog", 0, 14));
        label.setBorder(new Flush3DBorder());
        label.setForeground(java.awt.Color.black);
        table_panel.add(label);

        //�ݒ�l�e�[�u��
        c_old_table = new CtOldTable();
        JTableHeader tabHead = c_old_table.getTableHeader();
        tabHead.setReorderingAllowed(false);
        JScrollPane panel = new JScrollPane(c_old_table);
        panel.setBounds(10, 30, 178, 613);
        table_panel.add(panel);

        //�ύX�l�e�[�u��
        c_table = new CtTable();
        tabHead = c_table.getTableHeader();
        tabHead.setReorderingAllowed(false);
        panel = new JScrollPane(c_table);
        panel.setBounds(197, 30, 178, 613);
        table_panel.add(panel);

/*******************************************************************************/
        JButton input_button = new JButton("�{");
        input_button.setBounds(331, 655, 44, 24);
        input_button.setLocale(new Locale("ja","JP"));
        input_button.setFont(new java.awt.Font("dialog", 0, 24));
        input_button.setBorder(new Flush3DBorder());
        input_button.setForeground(java.awt.Color.black);
        input_button.addActionListener(new InputButton());
        table_panel.add(input_button);
/*******************************************************************************/

        JButton reset_button = new JButton("�ēǂݍ���");
        reset_button.setBounds(10, 685, 178, 24);
        reset_button.setLocale(new Locale("ja","JP"));
        reset_button.setFont(new java.awt.Font("dialog", 0, 18));
        reset_button.setBorder(new Flush3DBorder());
        reset_button.setForeground(java.awt.Color.black);
        reset_button.addActionListener(new ReLoadButton());
        table_panel.add(reset_button);

        JButton repaint_button = new JButton("��  �\  ��");
        repaint_button.setBounds(10, 745, 178, 24);
        repaint_button.setLocale(new Locale("ja","JP"));
        repaint_button.setFont(new java.awt.Font("dialog", 0, 18));
        repaint_button.setBorder(new Flush3DBorder());
        repaint_button.setForeground(java.awt.Color.black);
        repaint_button.addActionListener(new RepaintButton());
        table_panel.add(repaint_button);

        JButton del_button = new JButton("�I �� �� ��");
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

        JButton shift_down_button = new JButton("��");
        shift_down_button.setBounds(286, 685, 44, 24);
        shift_down_button.setLocale(new Locale("ja","JP"));
        shift_down_button.setFont(new java.awt.Font("dialog", 0, 18));
        shift_down_button.setBorder(new Flush3DBorder());
        shift_down_button.setForeground(java.awt.Color.black);
        shift_down_button.addActionListener(new ShiftDownButton());
        table_panel.add(shift_down_button);

        JButton shift_up_button = new JButton("��");
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

        JButton l_shift_down_button = new JButton("��");
        l_shift_down_button.setBounds(286, 715, 44, 24);
        l_shift_down_button.setLocale(new Locale("ja","JP"));
        l_shift_down_button.setFont(new java.awt.Font("dialog", 0, 18));
        l_shift_down_button.setBorder(new Flush3DBorder());
        l_shift_down_button.setForeground(java.awt.Color.black);
        l_shift_down_button.addActionListener(new l_ShiftDownButton());
        table_panel.add(l_shift_down_button);

        JButton l_shift_up_button = new JButton("��");
        l_shift_up_button.setBounds(331, 715, 44, 24);
        l_shift_up_button.setLocale(new Locale("ja","JP"));
        l_shift_up_button.setFont(new java.awt.Font("dialog", 0, 18));
        l_shift_up_button.setBorder(new Flush3DBorder());
        l_shift_up_button.setForeground(java.awt.Color.black);
        l_shift_up_button.addActionListener(new l_ShiftUpButton());
        table_panel.add(l_shift_up_button);
/**************20060529***************/


/**************20060529***************
        label = new JLabel("�k��",JLabel.CENTER);
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

        label = new JLabel("�q��",JLabel.CENTER);
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
        label = new JLabel("�k��",JLabel.CENTER);
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

        label = new JLabel("�q��",JLabel.CENTER);
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
			view_lab.setText("�}�X�^�[�\��");
			mst_show = true;
		}else{
			mstShowButton.setEnabled(true);
			view_lab.setText("�J�����g�\��");
			mst_show = false;
		}

        for(int i = 0 ; i < REC_MAX ; i++){
            c_old_table.setValueAt(null,i,1);
            c_old_table.setValueAt(null,i,2);
            c_table.setValueAt(null,i,1);
            c_table.setValueAt(null,i,2);
        }


        Vector dat =  CZSystem.getCtTb(edit_group,edit_recip,edit_number,edit_current);

        /* �}�X�^�[�f�[�^�擾 **************************************/
        Vector mstdat =  CZSystem.getCtTb(edit_group,edit_recip,edit_number,false);

        if(null == dat){
            current_data = null;
            if(edit_current){
                Object msg[] = {"���Ɛ���e�[�u��",
                                "�e�[�u�������݂��܂���I�I",
                                ""};
                errorMsg(msg);
            }
            else {
                Object msg[] = {"����e�[�u��",
                                "�e�[�u�������݂��܂���I�I",
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

//@@@ ��������@��ʂ̃T�C�Y������������B
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
//@@@ �����܂�
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
    //  �ύX�l�̂ݕ\�����t���b�V��
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

        // �e�[�u���S�ď��������̏���
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
    //  �ύX�l�̂ݕ\�����t���b�V��
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

        // �e�[�u���S�ď��������̏���
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

            // �Е��� null
            if(l == null && r != null) return false;
            if(l != null && r == null) return false;

            // ������ null
            if(null == l) continue;
            if(null == r) continue;

            CZSystemCtTb data = new CZSystemCtTb();
            data.l_val = l.floatValue();
            data.r_val = r.floatValue();

            // �k��
            if((edit_name.l_min > data.l_val) ||
                   (edit_name.l_max < data.l_val)){
                return false;
            }

            // �q��
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

        // �k���ɏ����\�[�g
        Arrays.sort(data, new Sort1());
        send_data = data;
        return true;
    }


    //
    // ���b�Z�[�W�̕\��
    //
    private boolean errorMsg(Object msg[]){
        JOptionPane.showMessageDialog(null,msg,
                                    "����e�[�u���G���[",
                                    JOptionPane.ERROR_MESSAGE);
        return true;
    }


    /*******************************************************
     *
     * �C���{�^���̏���
     *
     *******************************************************/
    class ModifyButton implements ActionListener {
        public void actionPerformed(ActionEvent ev){

            if(c_table.isEditing()){
                CZSystem.log("CZControlTableSub ModifyButton"," actionPerformed Table Data EDIT !!");
                Object msg[] = {"����e�[�u��",
                                "�ݒ蒆���ڗL��I�I",
                                ""};
                errorMsg(msg);
                return ;
            }

            if(1 > op_name.getText().length()){
                CZSystem.log("CZControlTableSub ModifyButton","actionPerformed Table Op Name Error !!");
                Object msg[] = {"����e�[�u��",
                                "�ݒ�҂���͂��Ă������I�I",
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
				// �ُ픻�菈��
				int jflg = 0;	// 0:�ُ픻�菈���s�v  1:�ُ픻�菈���J�n
				
				// �J�����g�e�[�u���Ǎ���
				Vector dat = null;
				if ((edit_group == 5) && (edit_number == 15)) {	// ��ʑ� ����No��15�̏ꍇ
					dat =  CZSystem.getCtTb(5,edit_recip,20,edit_current);		// �Ǎ��݃f�[�^�̍���No��20
					CZSystem.log("CZControlTableSub ","����No.20�̃f�[�^�T�C�Y:" + dat.size());
					jflg = 1;
				}else if ((edit_group == 5) && (edit_number == 20)) {	// ��ʑ� ����No��20�̏ꍇ
					dat =  CZSystem.getCtTb(5,edit_recip,15,edit_current);		// �Ǎ��݃f�[�^�̍���No��15
					CZSystem.log("CZControlTableSub ","����No.15�̃f�[�^�T�C�Y:" + dat.size());
					jflg = 1;
				}else{
					jflg = 0;	// �ُ픻�菈����~
				}
				
				if(jflg != 0){
					// ����Ώۃe�[�u���f�[�^�i�c�a�Ǎ��݃f�[�^�j����
					CZSystemCtTb data[] = new CZSystemCtTb[dat.size()];
					
					for(int i = 0 ; i < dat.size() ; i++){
						data[i] = (CZSystemCtTb)dat.elementAt(i);
						CZSystem.log("CZControlTableSub ","data.l_val[" + data[i].l_val + "] : data.r_val[" + data[i].r_val + "]");
					}
					
					CZSystem.log("CZControlTableSub ","��ʑ��@�k���ŏ��l : " + leftData[0]);
					CZSystem.log("CZControlTableSub ","��ʑ��@�k���ő�l : " + leftData[size-1]);
					CZSystem.log("CZControlTableSub ","�c�a���@�k���ŏ��l : " + data[0].l_val);
					CZSystem.log("CZControlTableSub ","�c�a���@�k���ő�l : " + data[dat.size()-1].l_val);
					
					// ��ʑ��Ƃc�a���̃f�[�^(�k���l)�����b�v���Ă��邩�`�F�b�N
					if(leftData[0] >= data[dat.size()-1].l_val){
						CZSystem.log("CZControlTableSub ","��ʑ��@�k���ŏ��l >= �c�a���@�k���ő�l");
						CZSystem.log("CZControlTableSub ","�J�����g�f�[�^�@�ُ픻�菈���s�v");
					}else if(leftData[size-1] <= data[0].l_val){
						CZSystem.log("CZControlTableSub ","��ʑ��@�k���ő�l <= �c�a���@�k���ŏ��l");
						CZSystem.log("CZControlTableSub ","�J�����g�f�[�^�@�ُ픻�菈���s�v");
					}else{
						CZSystem.log("CZControlTableSub ","�J�����g�f�[�^�@�ُ픻�菈���X�^�[�g");
						
						int lb_flg = 0;		// ���[�v�������f�t���O��1:���f
						for(int i = 0 ; i < size-1 ; i++){
							if(rightData[i] != rightData[i+1]){		// ��ʑ��@�q���l�ς̂k���͈͓���
								CZSystem.log("CZControlTableSub ","��ʑ�(L���l)   S [#" + (i+1) + "]: " + leftData[i] + " E [#" + (i+1+1) + "]: " + leftData[i+1]);
								CZSystem.log("CZControlTableSub ","��ʑ�(R���l)   S [#" + (i+1) + "]: " + rightData[i] + " E [#" + (i+1+1) + "]: " + rightData[i+1]);
								
								for(int j = 0 ; j < dat.size()-1 ; j++){
									if(data[j].r_val != data[j+1].r_val){	// �c�a���@�q���l�ς̂k���͈͓���
										CZSystem.log("CZControlTableSub ","�c�a��(L���l)   S [#" + (j+1) + "]: " + data[j].l_val + " E [#" + (j+1+1) + "]: " + data[j+1].l_val);
										CZSystem.log("CZControlTableSub ","�c�a��(R���l)   S [#" + (j+1) + "]: " + data[j].r_val + " E [#" + (j+1+1) + "]: " + data[j+1].r_val);
										
										// ��ʑ���(�k��)�ϔ͈͂ɂc�a����(�k��)�ϔ͈͂����b�v���Ă��邩�`�F�b�N
										if(leftData[i] >= data[j+1].l_val){
											CZSystem.log("CZControlTableSub ","�ϔ͈́@(��ʑ�)�k���ŏ��l " + leftData[i] + " >= " + "(�c�a��)�k���ő�l" + data[j+1].l_val);
											CZSystem.log("CZControlTableSub ","�ϔ͈̓��b�v�����I�@�ݒ�l�ُ햳���I");
										}else if(leftData[i+1] <= data[j].l_val){
											CZSystem.log("CZControlTableSub ","�ϔ͈́@(��ʑ�)�k���ő�l " + leftData[i+1] + " <= " + "(�c�a��)�k���ŏ��l" + data[j].l_val);
											CZSystem.log("CZControlTableSub ","�ϔ͈̓��b�v�����I�@�ݒ�l�ُ햳���I");
										}else{
											CZSystem.log("CZControlTableSub ","�ϔ͈̓��b�v����I�I�@�ݒ�l�ُ킠��I�I");
											
											Object msg[] = {"����ϐݒ�ُ�",
															"�}�O�l�b�g�P���ꋭ�xPF��",
															"�}�O�l�b�g�ʒuPF�ݒ��",
															"�m�F���Ă��������I"};
											errorMsg(msg);
											lb_flg = 1;		//���[�v�����I���t���O�Z�b�g
										}
									}
									
									if(lb_flg == 1){	//���[�v�����I��
										CZSystem.log("CZControlTableSub ","���[�v�����I��");
										break;
									}
								}
							}
							
							if(lb_flg == 1){	//���[�v�����I��
								CZSystem.log("CZControlTableSub ","���[�v�����I��");
								break;
							}
						}
					}
				}else{
					CZSystem.log("CZControlTableSub ","�J�����g�f�[�^�@�ُ픻�菈���s�v");
				}
				
// @@@@@ 2014.01.09

                CZSystem.CZControlTableExchange(op_name.getText(),MODIFY_DATA,edit_group,
                                                edit_recip, edit_number,leftData,rightData);

            }
            else {
                Object msg[] = {"����e�[�u��",
                                "�l���m�F���Ă������I�I",
                                ""};
                errorMsg(msg);
            }
            return ;
        }
    }

    /*******************************************************
     *
     * �ۑ��{�^��
     *
     *******************************************************/
    class SaveButton implements ActionListener {
        public void actionPerformed(ActionEvent ev){

            if(c_table.isEditing()){
                CZSystem.log("CZControlTableSub SaveButton","actionPerformed Table Data EDIT !!");
                Object msg[] = {"����e�[�u��",
                                "�ݒ蒆���ڗL��I�I",
                                ""};
                errorMsg(msg);
                return ;
            }

            if(1 > op_name.getText().length()){
                CZSystem.log("CZControlTableSub SaveButton","actionPerformed Table Op Name Error !!");

                Object msg[] = {"����e�[�u��",
                                "�ݒ�҂���͂��Ă��������I�I",
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
				// �ُ픻�菈��
				int jflg = 0;	// 0:�ُ픻�菈���s�v  1:�ُ픻�菈���J�n
				
				// �}�X�^�[�e�[�u���Ǎ���
				Vector dat = null;
				if ((edit_group == 5) && (edit_number == 15)) {	// ��ʑ� ����No��15�̏ꍇ
					dat =  CZSystem.getCtTb(5,edit_recip,20,false);		// �Ǎ��݃f�[�^�̍���No��20
					CZSystem.log("CZControlTableSub ","����No.20�̃f�[�^�T�C�Y:" + dat.size());
					jflg = 1;
				}else if ((edit_group ==5) && (edit_number == 20)) {	// ��ʑ� ����No��20�̏ꍇ
					dat =  CZSystem.getCtTb(5,edit_recip,15,false);		// �Ǎ��݃f�[�^�̍���No��15
					CZSystem.log("CZControlTableSub ","����No.15�̃f�[�^�T�C�Y:" + dat.size());
					jflg = 1;
				}else{
					jflg = 0;	// �ُ픻�菈����~
				}
				
				if(jflg != 0){
					// ����Ώۃe�[�u���f�[�^�i�c�a�Ǎ��݃f�[�^�j����
					CZSystemCtTb data[] = new CZSystemCtTb[dat.size()];
					
					for(int i = 0 ; i < dat.size() ; i++){
						data[i] = (CZSystemCtTb)dat.elementAt(i);
						CZSystem.log("CZControlTableSub ","data.l_val[" + data[i].l_val + "] : data.r_val[" + data[i].r_val + "]");
					}
					
					CZSystem.log("CZControlTableSub ","��ʑ��@�k���ŏ��l : " + leftData[0]);
					CZSystem.log("CZControlTableSub ","��ʑ��@�k���ő�l : " + leftData[size-1]);
					CZSystem.log("CZControlTableSub ","�c�a���@�k���ŏ��l : " + data[0].l_val);
					CZSystem.log("CZControlTableSub ","�c�a���@�k���ő�l : " + data[dat.size()-1].l_val);
					
					// ��ʑ��Ƃc�a���̃f�[�^(�k���l)�����b�v���Ă��邩�`�F�b�N
					if(leftData[0] >= data[dat.size()-1].l_val){
						CZSystem.log("CZControlTableSub ","��ʑ��@�k���ŏ��l >= �c�a���@�k���ő�l");
						CZSystem.log("CZControlTableSub ","�J�����g�f�[�^�@�ُ픻�菈���s�v");
					}else if(leftData[size-1] <= data[0].l_val){
						CZSystem.log("CZControlTableSub ","��ʑ��@�k���ő�l <= �c�a���@�k���ŏ��l");
						CZSystem.log("CZControlTableSub ","�}�X�^�[�f�[�^�@�ُ픻�菈���s�v");
					}else{
						CZSystem.log("CZControlTableSub ","�}�X�^�[�f�[�^�@�ُ픻�菈���X�^�[�g");
						
						int lb_flg = 0;		// ���[�v�������f�t���O��1:���f
						for(int i = 0 ; i < size-1 ; i++){
							if(rightData[i] != rightData[i+1]){		// ��ʑ��@�q���l�ς̂k���͈͓���
								CZSystem.log("CZControlTableSub ","��ʑ�(L���l)   S [#" + (i+1) + "]: " + leftData[i] + " E [#" + (i+1+1) + "]: " + leftData[i+1]);
								CZSystem.log("CZControlTableSub ","��ʑ�(R���l)   S [#" + (i+1) + "]: " + rightData[i] + " E [#" + (i+1+1) + "]: " + rightData[i+1]);
								
								for(int j = 0 ; j < dat.size()-1 ; j++){
									if(data[j].r_val != data[j+1].r_val){	// �c�a���@�q���l�ς̂k���͈͓���
										CZSystem.log("CZControlTableSub ","�c�a��(L���l)   S [#" + (j+1) + "]: " + data[j].l_val + " E [#" + (j+1+1) + "]: " + data[j+1].l_val);
										CZSystem.log("CZControlTableSub ","�c�a��(R���l)   S [#" + (j+1) + "]: " + data[j].r_val + " E [#" + (j+1+1) + "]: " + data[j+1].r_val);
										
										// ��ʑ���(�k��)�ϔ͈͂ɂc�a����(�k��)�ϔ͈͂����b�v���Ă��邩�`�F�b�N
										if(leftData[i] >= data[j+1].l_val){
											CZSystem.log("CZControlTableSub ","�ϔ͈́@(��ʑ�)�k���ŏ��l " + leftData[i] + " >= " + "(�c�a��)�k���ő�l" + data[j+1].l_val);
											CZSystem.log("CZControlTableSub ","�ϔ͈̓��b�v�����I�@�ݒ�l�ُ햳���I");
										}else if(leftData[i+1] <= data[j].l_val){
											CZSystem.log("CZControlTableSub ","�ϔ͈́@(��ʑ�)�k���ő�l " + leftData[i+1] + " <= " + "(�c�a��)�k���ŏ��l" + data[j].l_val);
											CZSystem.log("CZControlTableSub ","�ϔ͈̓��b�v�����I�@�ݒ�l�ُ햳���I");
										}else{
											CZSystem.log("CZControlTableSub ","�ϔ͈̓��b�v����I�I�@�ݒ�l�ُ킠��I�I");
											
											Object msg[] = {"����ϐݒ�ُ�",
															"�}�O�l�b�g�P���ꋭ�xPF��",
															"�}�O�l�b�g�ʒuPF�ݒ��",
															"�m�F���Ă��������I"};
											errorMsg(msg);
											lb_flg = 1;		//���[�v�����I���t���O�Z�b�g
										}
									}
									
									if(lb_flg == 1){	//���[�v�����I��
										CZSystem.log("CZControlTableSub ","���[�v�����I��");
										break;
									}
								}
							}
							
							if(lb_flg == 1){	//���[�v�����I��
								CZSystem.log("CZControlTableSub ","���[�v�����I��");
								break;
							}
						}
					}
				}else{
					CZSystem.log("CZControlTableSub ","�}�X�^�[�f�[�^�@�ُ픻�菈���s�v");
				}
				
// @@@@@ 2014.01.09

                CZSystem.CZControlTableExchange(op_name.getText(),SAVE_DATA,edit_group,
                                                edit_recip, edit_number,leftData,rightData);
            }
            else {
                Object msg[] = {"����e�[�u��",
                                "�l���m�F���Ă��������I�I",
                                ""};
                errorMsg(msg);
            }
            return ;
        }
    }


    /*******************************************************
     *
     * �I���{�^��
     *
     *******************************************************/
    class CancelButton implements ActionListener {
        public void actionPerformed(ActionEvent ev){
            setVisible(false);
        }
    }


    /*******************************************************
     *
     * �C���l���̓{�^���̏���
     *
     *******************************************************/
    class InputButton implements ActionListener {
        public void actionPerformed(ActionEvent ev){
			Float l_val1; // ���͒l�iL���j
			Float r_val1; // �v�Z���ʒl�iR���j
			Float l_val2; // �i�������j���͒l���ЂƂ������l�iL���j
			Float r_val2; // �i�������j���͒l���ЂƂ������l�iR���j
			Float l_val3; // �i�������j���͒l���ЂƂ傫���l�iL���j
			Float r_val3; // �i�������j���͒l���ЂƂ傫���l�iR���j
			Float l_val4; // �i�~�����j���͒l���ЂƂ傫���l�iL���j
			Float r_val4; // �i�~�����j���͒l���ЂƂ傫���l�iR���j
			Float l_val5; // �i�~�����j���͒l���ЂƂ������l�iL���j
			Float r_val5; // �i�~�����j���͒l���ЂƂ������l�iR���j
			
			
			CZSystem.log("CZControlTableSub InputButton","edit_name.k_sort: " + edit_name.k_sort);
			
			// �I���s
			int row = c_table.getSelectedRow();
			
			if(row == -1){
				return;
			}
			
			// ���͒l�iL���j
			l_val1 = (Float)c_table.getValueAt(row,1);
			
			if(l_val1 == null){
				return;
			}
			
			if(edit_name.k_sort == 1){	/* �l�������̂Ƃ� */
				for(int i = 0; i < REC_MAX; i++){
					l_val3 = (Float)c_table.getValueAt(i,1);
					CZSystem.log("CZControlTableSub InputButton","l_val3 ���͒l���ЂƂ傫���l�iL���j: " + l_val3);
					
					if(l_val3 == null){
						Object msg[] = {"���͂������l���Q�_�Ԃ̒l�ł͂���܂���",
										"�l���m�F���Ă��������I�I",
										""};
						errorMsg(msg);
						return;
					}
					
					if(l_val1 < l_val3){
						if(i == 0){
							Object msg[] = {"���͂������l���Q�_�Ԃ̒l�ł͂���܂���",
											"�l���m�F���Ă��������I�I",
											""};
							errorMsg(msg);
							return;
						}
						
						// �Q�_�Ԃ̒l���擾
						l_val2 = (Float)c_table.getValueAt(i-1,1);
						r_val2 = (Float)c_table.getValueAt(i-1,2);
						l_val3 = (Float)c_table.getValueAt(i,1);
						r_val3 = (Float)c_table.getValueAt(i,2);
						
						CZSystem.log("CZControlTableSub InputButton","l_val1�i���͒l�j : " + l_val1);
						CZSystem.log("CZControlTableSub InputButton","l_val2 ���͒l���ЂƂ������l�iL���j: " + l_val2);
						CZSystem.log("CZControlTableSub InputButton","r_val2 ���͒l���ЂƂ������l�iR���j: " + r_val2);
						CZSystem.log("CZControlTableSub InputButton","l_val3 ���͒l���ЂƂ傫���l�iL���j: " + l_val3);
						CZSystem.log("CZControlTableSub InputButton","r_val3 ���͒l���ЂƂ傫���l�iR���j: " + r_val3);
						
						// �v�Z���ʒl�iR���j
						r_val1 = ((r_val3 - r_val2) / (l_val3 - l_val2) * (l_val1 - l_val2) + r_val2);
						
						CZSystem.log("CZControlTableSub InputButton","r_val1�i�v�Z���ʒl�iR���j�j: " + r_val1);
						
						if(true == l_val1.equals(l_val2)){
							Object msg[] = {"���͂������l���Q�_�Ԃ̒l�ł͂���܂���",
											"�l���m�F���Ă��������I�I",
											""};
							errorMsg(msg);
							return;
						}
						
						// �����v�Z���ʒl��\��
						c_table.setValueAt(l_val1,row,1);
						c_table.setValueAt(r_val1,row,2);
						
						// �s�I������
						c_table.clearSelection();
						return;
					}
				}
			}else{	/* �l���~���̂Ƃ� */
				for(int i = 0; i < REC_MAX; i++){
					l_val4 = (Float)c_table.getValueAt(i,1);
					CZSystem.log("CZControlTableSub InputButton","l_val4 ���͒l���ЂƂ傫���l�iL���j: " + l_val4);
					
					if(l_val4 == null){
						Object msg[] = {"���͂������l���Q�_�Ԃ̒l�ł͂���܂���",
										"�l���m�F���Ă��������I�I",
										""};
						errorMsg(msg);
						return;
					}
					
					if(l_val1 > l_val4){
						if(i == 0){
							Object msg[] = {"���͂������l���Q�_�Ԃ̒l�ł͂���܂���",
											"�l���m�F���Ă��������I�I",
											""};
							errorMsg(msg);
							return;
						}
						
						// �Q�_�Ԃ̒l���擾
						l_val4 = (Float)c_table.getValueAt(i-1,1);
						r_val4 = (Float)c_table.getValueAt(i-1,2);
						l_val5 = (Float)c_table.getValueAt(i,1);
						r_val5 = (Float)c_table.getValueAt(i,2);
						
						CZSystem.log("CZControlTableSub InputButton","l_val1�i���͒l�j : " + l_val1);
						CZSystem.log("CZControlTableSub InputButton","l_val4 ���͒l���ЂƂ傫���l�iL���j: " + l_val4);
						CZSystem.log("CZControlTableSub InputButton","r_val4 ���͒l���ЂƂ傫���l�iR���j: " + r_val4);
						CZSystem.log("CZControlTableSub InputButton","l_val5 ���͒l���ЂƂ������l�iL���j: " + l_val5);
						CZSystem.log("CZControlTableSub InputButton","r_val5 ���͒l���ЂƂ������l�iR���j: " + r_val5);
						
						// �v�Z���ʒl�iR���j
						r_val1 = ((r_val4 - r_val5) / (l_val4 - l_val5) * (l_val1 - l_val5) + r_val5);
						
						CZSystem.log("CZControlTableSub InputButton","r_val1�i�v�Z���ʒl�iR���j�j: " + r_val1);
						
						if(true == l_val1.equals(l_val4)){
							Object msg[] = {"���͂������l���Q�_�Ԃ̒l�ł͂���܂���",
											"�l���m�F���Ă��������I�I",
											""};
							errorMsg(msg);
							return;
						}
						
						// �����v�Z���ʒl��\��
						c_table.setValueAt(l_val1,row,1);
						c_table.setValueAt(r_val1,row,2);
						
						// �s�I������
						c_table.clearSelection();
						return;
					}
				}
			}
        }
    }

    /*******************************************************
     *
     * �ēǍ��{�^���̏���
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
     * �ĕ\��
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
     * �I���폜�{�^���̏���
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
     * ���{�^���̏���
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
     * ���{�^���̏���
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
     * ���{�^���̏��� 20060529
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
     * ���{�^���̏��� 20060529
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
     * �}�X�^�[�\���{�^�� @@@@
     *
     *******************************************************/
    class MasterShowAction implements ActionListener {
        public void actionPerformed(ActionEvent ev){

            if (mst_show == true) {
				view_lab.setText("�J�����g�\��");
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
				view_lab.setText("�}�X�^�[�\��");
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
     * �k���{�^�� @@@
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
     * ��{�^�� @@@
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
     * �g��{�^�� @@@
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
     *   �ݒ�l�F����e�[�u��
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

                // ����No
                colum = cmdl.getColumn(0);
                colum.setMaxWidth(40);
                colum.setMinWidth(40);
                colum.setWidth(40);

                // �k��
                colum = cmdl.getColumn(1);
                colum.setMaxWidth(60);
                colum.setMinWidth(60);
                colum.setWidth(60);
                render1 = new CtOldTblRenderer();
                colum.setCellRenderer(render1);

                // �q��
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
         *       �ݒ�l�F����e�[�u���F���f��
         *
         */
        public class CtOldTblMdl extends AbstractTableModel {

            private int TBL_ROW = REC_MAX;
            final   int TBL_COL = 3;

            final String[] names = {" # " , "�k��" , "�q��"};

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
         *   �ݒ�l�F����e�[�u���F�����_�[
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

                // �\���t�H�[�}�b�g
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
				// �F�ύX
				//
				public void setColor(Color col){

					fColor = col;
				}

            } // CtOldTblRenderer
        } // CtOldTable

        /***************************************************
         *
         *   �ύX�l�F����e�[�u��
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

                    // ����No
                    colum = cmdl.getColumn(0);
                    colum.setMaxWidth(40);
                    colum.setMinWidth(40);
                    colum.setWidth(40);

                    // �k��
                    colum = cmdl.getColumn(1);
                    colum.setMaxWidth(60);
                    colum.setMinWidth(60);
                    colum.setWidth(60);
                    colum.setCellRenderer(new CtTblRenderer());

                    // �q��
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
             *       �ύX�l�F����e�[�u���F���f��
             *
             ***********************************************/
            public class CtTblMdl extends AbstractTableModel {

                private int TBL_ROW = REC_MAX;
                final   int TBL_COL = 3;

                final String[] names = {" # " , "�k��" , "�q��"};

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
         *   �ύX�l�F����e�[�u���F�����_�[
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

                // �\���t�H�[�}�b�g
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
     *  L�l�������Ń\�[�g�������̂̏ꍇ�A�P�i�ʏ�̂��́j
     *  L�l���~���Ń\�[�g�������̂̏ꍇ�A�Q�i�c�t���j
     *
     *   �����Ń\�[�g
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
     * �O���t���C���p�l���̃X�N���[��
     *
     *******************************************************/
    class MainPanel extends JScrollPane {

        // ���
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
        * ���X�i�[��ǉ�����B
        */
        public void addComponentListener( ComponentListener listener ) {
            view_.addComponentListener( listener );
        }
        /**
        * ��ʎ擾
        * @param
        * @return
        */
        public MainPanelView getView(){
            return view_;
        }

        /**
        * ��ʈʒu�擾
        * @param
        * @return
        */
        public Point getViewLocation(){
            return view_.getLocation();
        }

        /**
        * ��ʈʒu�ݒ�
        * @param
        * @return
        */
        public void setViewLocation( Point pos ){
            view_.setLocation( pos );
        }
        /**
        * ��ʕ��ݒ�
        * @param
        * @return
        */
        public void setViewSize( int width, int height){
            view_.setMainPanelViewSize( width, height );
        }

        /**
        * �`��
        */
        public void viewPaint(){
            view_.repaint();
        }
    } // MainPanel

    /*******************************************************
     *
     * �O���t���C���p�l��
     *
     *******************************************************/
    class MainPanelView extends JPanel {

        /**
        * �R���X�g���N�^
        */
        MainPanelView(){
            super();
            setLayout(null);
            setBackground(java.awt.Color.white);
        }

        /**
        * ��ʂ̃T�C�Y��ݒ肷��B
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

            // �O���t�ڐ��̕`��
            g.setColor(java.awt.Color.darkGray);
            // �w��
            float bun = (float)d.width / (l_graph_bun * 5);
            float x  = 0;
            for(int j = 0 ; j < (l_graph_bun * 5)+1 ; j++){
                g.drawLine((int)x,0,(int)x,d.height);
                x+=bun;
            }

            // �x��
            bun = (float)d.height / (r_graph_bun * 5);
            float y   = 0.0f;
            for(int i = 0 ; i < (r_graph_bun * 5)+1 ; i++){
                g.drawLine(0,(int)y,d.width,(int)y);
                y+=bun;
            }

            g.setColor(java.awt.Color.lightGray);
            // �w��
            bun = (float)d.width / (float)l_graph_bun;
            x   = 0.0f;
            for(int i = 0 ; i < l_graph_bun+1 ; i++){
                g.drawLine((int)x,0,(int)x,d.height);
                x+=bun;
            }

            // �x��
            bun = (float)d.height / (float)r_graph_bun;
            y   = 0.0f;
            for(int i = 0 ; i < r_graph_bun+1 ; i++){
                g.drawLine(0,(int)y,d.width,(int)y);
                y+=bun;
            }

            if(null == current_data) return;

            int   size = current_data.size();
            if(2 > size) return;

            //�O���t���f�[�^�̃\�[�g
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

            //�O���t�̕`��
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

            //�O���t���f�[�^�̃\�[�g
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
        // �ݒ�f�[�^�̕`��
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
            //�O���t���f�[�^�̃\�[�g
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

            //�O���t�̕`��
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
//2003.11.27new�̐��̑������݂�ȂƂ��킷���߃R�����g
            //���݂̐ݒ��st_tmp�Ɋi�[
//2003.11.27            Stroke st_tmp = g2.getStroke();
			//���C���̑������s�N�Z���Őݒ�i2f�Œʏ탉�C���̔{�̑����j
//2003.11.27            BasicStroke bs = new BasicStroke(2f);  // 10�s�N�Z����
			//���C���̑�����ύX�����ݒ���Z�b�g����
//2003.11.27            g2.setStroke(bs);

			//���C���̐F��ύX����
            g2.setColor(NEW_PRO_COL);
			//���C����\���i���j
            g2.drawPolyline(new_val_x,new_val_y,size);
			//���̐ݒ�ɖ߂�
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
            //�O���t���f�[�^�̃\�[�g
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

            //�O���t�̕`��
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
//2003.11.27new�̐��̑������݂�ȂƂ��킷���߃R�����g
            //���݂̐ݒ��st_tmp�Ɋi�[
//2003.11.27            Stroke st_tmp = g2.getStroke();
			//���C���̑������s�N�Z���Őݒ�i2f�Œʏ탉�C���̔{�̑����j
//2003.11.27            BasicStroke bs = new BasicStroke(2f);  // 10�s�N�Z����
			//���C���̑�����ύX�����ݒ���Z�b�g����
//2003.11.27            g2.setStroke(bs);

			//���C���̐F��ύX����
            g2.setColor(java.awt.Color.cyan);
			//���C����\���i���j
            g2.drawPolyline(new_val_x,new_val_y,size);
			//���̐ݒ�ɖ߂�
//2003.11.27          g2.setStroke(st_tmp);
        }
/**********************************************/


		@SuppressWarnings("unchecked")
        private void repaint_mst_data(Graphics g,Dimension d){

            int size = master_data.size();
            //�O���t���f�[�^�̃\�[�g
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

            //�O���t�̕`��
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

			//���C���̐F��ύX����
			g2.setColor(java.awt.Color.yellow);
			//���C����\���i���j
			g2.drawPolyline(new_val_x,new_val_y,size);
			//���̐ݒ�ɖ߂�
        }


        //
        // �o�u�f�[�^�̕`��
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
        * �o�u�f�[�^�̕`�� �����グ�v���t�@�C��
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
        // �o�u�f�[�^�̕`�� �{�f�B�[�̃v���t�@�C���̏ꍇ
        //
        private void repaint_pv_data_body_pf(Graphics g,Dimension d,int val_no,int pf_no,boolean shift_flg){

//@@            CZSystem.log("CZControlTableSub repaint_pv_data_body_pf",
//@@                    "GROUP[" + edit_group + "] NUMBER[" + edit_number + "]");

            int size = pv_data_body.size();
            if(1 > size) return;

            //�O���t�̕`��
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

            // ���уv���t�@�C��
            g.setColor(VAL_PRO_COL);
            g.drawPolyline(new_val_x,new_prf_y,size);

            // ����
            g.setColor(VAL_COL);
            g.drawPolyline(new_val_x,new_val_y,size);
        }
    } // MainPanelView

    /*******************************************************
     *
     * Y���̖ڐ��p�l���̃X�N���[��
     *
     *******************************************************/
    class RPanel extends JScrollPane {

        // ���
        private RPanelView view_;

        /**
        * �R���X�g���N�^
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
        * ���X�i�[��ǉ�����B
        */
        public void addComponentListener( ComponentListener listener ) {
            view_.addComponentListener( listener );
        }
        /**
        * ��ʎ擾
        * @param
        * @return
        */
        public RPanelView getView(){
            return view_;
        }

        /**
        * ��ʈʒu�擾
        * @param
        * @return
        */
        public Point getViewLocation(){
            return view_.getLocation();
        }

        /**
        * ��ʈʒu�ݒ�
        * @param
        * @return
        */
        public void setViewLocation( Point pos ){
            view_.setLocation( pos );
        }
        /**
        * ��ʕ��ݒ�
        * @param
        * @return
        */
        public void setViewSize( int height){
            view_.setRPanelViewSize( height );
        }
    } // RPanel

    /*******************************************************
     *
     * Y���̖ڐ��p�l��
     *
     *******************************************************/
    class RPanelView extends JPanel {

        /**
        * �R���X�g���N�^
        */
        RPanelView(){
            super();
            setFont(new java.awt.Font("dialog", 0, 12));
        }
        /**
        * ��ʂ̃T�C�Y��ݒ肷��B
        */
        public void setRPanelViewSize( int height ) {
            setPreferredSize( new Dimension( 50, height ) );
            setSize( new Dimension( 50, height) );
            repaint();
        }

        /**
        * �`��
        */
        public void paint(Graphics g){

            Dimension d = getSize(null);
            g.setColor(java.awt.Color.black);
            g.fillRect(0,0,d.width,d.height);

            // �O���t�ڐ��̕`��
            g.setColor(java.awt.Color.darkGray);

            float bun = (float)d.height / (float)r_graph_bun;
            float y   = 0;
            for(int i = 0 ; i < r_graph_bun+1 ; i++){
                g.drawLine(0,(int)y,d.width,(int)y);
                y+=bun;
            }


            //�\���t�H�[�}�b�g
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
     * X���̖ڐ��p�l���̃X�N���[��
     *
     *******************************************************/
    class LPanel extends JScrollPane {

        // ���
        private LPanelView view_;

        /**
        * �R���X�g���N�^
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
        * ���X�i�[��ǉ�����B
        */
        public void addComponentListener( ComponentListener listener ) {
            view_.addComponentListener( listener );
        }
        /**
        * ��ʎ擾
        * @param
        * @return
        */
        public LPanelView getView(){
            return view_;
        }

        /**
        * ��ʈʒu�擾
        * @param
        * @return
        */
        public Point getViewLocation(){
            return view_.getLocation();
        }

        /**
        * ��ʈʒu�ݒ�
        * @param
        * @return
        */
        public void setViewLocation( Point pos ){
            view_.setLocation( pos );
        }
        /**
        * ��ʕ��ݒ�
        * @param
        * @return
        */
        public void setViewSize( int width){
            view_.setLPanelViewSize( width );
        }

    } // LPanel

    /*******************************************************
     *
     * X���̖ڐ��p�l��
     *
     *******************************************************/
    class LPanelView extends JPanel {

        /**
        * �R���X�g���N�^
        */
        LPanelView(){
            super();
            setFont(new java.awt.Font("dialog", 0, 12));
        }

        /**
        * ��ʂ̃T�C�Y��ݒ肷��B
        */
        public void setLPanelViewSize( int width ) {
            setPreferredSize( new Dimension( width, 50 ) );
            setSize( new Dimension( width, 50) );
            repaint();
        }

        /**
        * �`��
        */
        public void paint(Graphics g){

            Dimension d = getSize(null);
            g.setColor(java.awt.Color.black);
            g.fillRect(0,0,d.width,d.height);

            // �O���t�ڐ��̕`��
            g.setColor(java.awt.Color.darkGray);
            float bun = (float)d.width / (float)l_graph_bun;
            float x   = 0;
            for(int i = 0 ; i < l_graph_bun ; i++){
                g.drawLine((int)x,0,(int)x,d.height);
                x+=bun;
            }

            //�\���t�H�[�}�b�g
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
     *       �ݒ�҂���͂���TextField
     *
     ***************************************************************************/
    /*public*/ class TText extends JTextField {

        /**
        * �R���X�g���N�^
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
     *       �ݒ�l�V�t�g�ʂ���͂���TextField
     *
     ***************************************************************************/
    /*public*/ class ShiftText extends JTextField {

        /**
        * �R���X�g���N�^
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
     *       �O���t����������͂���TextField
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
     * �O���t�̃X�N���[���R���|�[�l���g���X�i�[ @@@
     *
     *******************************************************/
    private class GraphComponentListener implements ComponentListener {
        /**
        *�@�R���X�g���N�^
        */
        GraphComponentListener(){
            super();
        }

        /**
        * �X�N���[�����ړ��������̏���
        */
        public void componentMoved( java.awt.event.ComponentEvent e )
        {
            if( main_panelView == e.getComponent() ) {
                // ���C����ʂ��ړ������Ƃ���Y���ڐ����ړ�����
                Point mainViewPos = main_panelView.getLocation();
                Point yViewPos = r_panelView.getLocation();
                yViewPos.y = mainViewPos.y;
                r_panelView.setLocation( yViewPos );
            }
            else if( l_panelView == e.getComponent() ) {
                // X����ʂ��ړ��������̓��C����ʂ��ړ�����B
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
