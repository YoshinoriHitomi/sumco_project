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
 *   �s�o�f��{�O���t
 * @author  (KPK Co.,Ltd.)
 * @version 1.0 (2003/04/01)
 * @@ T6 ... �ǉ�
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


    private final int MAIN1_H_T     = 14;   // 15   ���C���q�[�^�[�P���x
    private final int MAIN1_H_T_PF  = 66;   // 67   ���C���q�[�^�[�P���x�v���t�@�C��
    private final int DIA           = 24;   // 25   ���a
    private final int DIA_PF        = 23;   // 24   ���a�v���t�@�C��
    private final int SXL_ST        = 17;   // 18   �����グ���x
    private final int SXL_ST_PF     = 75;   // 76   �����グ���x�v���t�@�C��
    private final int SXL_RT        = 18;   // 19   �V�[�h��]
    private final int SXL_RT_PF     = 80;   // 81   �V�[�h��]�v���t�@�C��
    private final int CRU_RT        = 20;   // 21   ���c�{��]
    private final int CRU_RT_PF     = 86;   // 87   ���c�{��]�v���t�@�C��
    private final int PULL_AR       = 15;   // 16   �v���A���S��
    private final int PULL_AR_PF    = 71;   // 72   �v���A���S���v���t�@�C��
    private final int VAC           = 32;   // 33   �F����
    private final int VAC_PF        = 88;   // 89   �F�����v���t�@�C��

    private String  main1_h_t_min_pro;      //���C���q�[�^�[�P���x
    private String  main1_h_t_max_pro;
    private String  main1_h_t_pf_min_pro;   //���C���q�[�^�[�P���x�v���t�@�C��
    private String  main1_h_t_pf_max_pro;
    private String  dia_min_pro;            //���a
    private String  dia_max_pro;
    private String  dia_pf_min_pro;         //���a�v���t�@�C��
    private String  dia_pf_max_pro;
    private String  sxl_st_min_pro;         //�����グ���x
    private String  sxl_st_max_pro;
    private String  sxl_st_pf_min_pro;      //�����グ���x�v���t�@�C��
    private String  sxl_st_pf_max_pro;
    private String  sxl_rt_pf_min_pro;      //�V�[�h��]�v���t�@�C��
    private String  sxl_rt_pf_max_pro;
    private String  cru_rt_pf_min_pro;      //���c�{��]�v���t�@�C��
    private String  cru_rt_pf_max_pro;
    private String  pull_ar_pf_min_pro;     //�v���A���S���v���t�@�C��
    private String  pull_ar_pf_max_pro;
    private String  vac_pf_min_pro;         //�F�����v���t�@�C��
    private String  vac_pf_max_pro;

    public float    x_length_mouse          = 0.0f;     //�r�w�k����
    public float    y_main1_h_t_mouse       = 0.0f;     //���C���q�[�^�[�P���x
    public float    y_main1_h_t_pf_mouse    = 0.0f;     //���C���q�[�^�[�P���x�v���t�@�C��
    public float    y_dia_mouse             = 0.0f;     //���a
    public float    y_sxl_st_mouse          = 0.0f;     //�����グ���x
    public float    y_sxl_st_pf_mouse       = 0.0f;     //�����グ���x�v���t�@�C��

    private final String GR_X_LENGTH_DEF    = "2500";   //�w���̒���
    private String  gr_x_length = GR_X_LENGTH_DEF;      //�w���̒���
    private float   gr_x_bun    = 10.0f;                //�w���̕���
    private float   gr_y_bun    = 5.0f;                 //�x���̕���

    private int Y_VIEW_TIMES    = 2;                    //�x���̔{��

    private String  X_LENGTH_LIST[] = {gr_x_length,
                       "2000",
                       "1500",
                       "1000",
                       "500",
                       "250",
                       "200",
                       "100",
                       "50"};

    private String ro_name                  = null; //�ΏۘF��
    private String ro_db_name               = null; //�ΏۘF�f�[�^�x�[�X��

    private CZSystemStart ro_bt_start       = null; //�����p�����グ����
    private Vector ro_bt_all_condition      = null; //�SBt�̈����グ����

    private Vector pv_data_shld             = null; //�V�����_�[�̃f�[�^
    private Vector pv_data_body             = null; //�{�f�B�[�̃f�[�^

    private JLabel main_ro_name_lab         = null; //�F�ԕ\��

    private MainSc  main_sc                 = null; //���C���O���t�X�N���[���p�l��
    private XSc     x_sc                    = null; //�w���O���t�X�N���[���p�l��
    private Y1Sc    y1_sc                   = null; //�x�������O���t�X�N���[���p�l��
    private Y2Sc    y2_sc                   = null; //�x���E���O���t�X�N���[���p�l��

    private MainMouseView   main_mouse_view = null; //�}�E�X���W�\���p�l��

    private SercheDialog    serche_dia      = null; //�����p�_�C�A���O
    private YLengDialog     y_leng_dia      = null; //�x���ݒ�p�_�C�A���O

    private ConditionPanel  conpane         = null; //�����A�����グ�����p�l��
    private GraphPanel      grapane         = null; //�O���t�ݒ�p�l��
    private SimplGraphPanel simpgrapane     = null; //�ȈՃO���t�\���p�l��

    // ---------- �R���X�g���N�^ ---------------------------
    //
    CZTPGMain(){
        super();

        try{
            // ----- Propertie_File��� Min,Max�l���擾����B --------
            Properties prop =  new Properties();
            FileInputStream pros = new FileInputStream(CZSystemDefine.TPGPROPERTY_FILE);
            prop.load(pros);

            prop.list(System.out);
            main1_h_t_min_pro       = prop.getProperty("MAIN1_H_T_MIN");    //���C���q�[�^�[�P���x
            main1_h_t_max_pro       = prop.getProperty("MAIN1_H_T_MAX");
            main1_h_t_pf_min_pro    = prop.getProperty("MAIN1_H_T_PF_MIN"); //���C���q�[�^�[�P���x�v���t�@�C��
            main1_h_t_pf_max_pro    = prop.getProperty("MAIN1_H_T_PF_MAX");
            dia_min_pro         = prop.getProperty("DIA_MIN");              //���a
            dia_max_pro         = prop.getProperty("DIA_MAX");
            dia_pf_min_pro      = prop.getProperty("DIA_PF_MIN");           //���a�v���t�@�C��
            dia_pf_max_pro      = prop.getProperty("DIA_PF_MAX");

            sxl_st_min_pro      = prop.getProperty("SXL_ST_MIN");           //�����グ���x�v���t�@�C��
            sxl_st_max_pro      = prop.getProperty("SXL_ST_MAX");
            sxl_st_pf_min_pro   = prop.getProperty("SXL_ST_PF_MIN");        //�����グ���x�v���t�@�C��
            sxl_st_pf_max_pro   = prop.getProperty("SXL_ST_PF_MAX");
            sxl_rt_pf_min_pro   = prop.getProperty("SXL_RT_PF_MIN");        //�V�[�h��]�v���t�@�C��
            sxl_rt_pf_max_pro   = prop.getProperty("SXL_RT_PF_MAX");
            cru_rt_pf_min_pro   = prop.getProperty("CRU_RT_PF_MIN");        //���c�{��]�v���t�@�C��
            cru_rt_pf_max_pro   = prop.getProperty("CRU_RT_PF_MAX");
            pull_ar_pf_min_pro  = prop.getProperty("PULL_AR_PF_MIN");       //�v���A���S���v���t�@�C��
            pull_ar_pf_max_pro  = prop.getProperty("PULL_AR_PF_MAX");
            vac_pf_min_pro      = prop.getProperty("VAC_PF_MIN");           //�F�����v���t�@�C��
            vac_pf_max_pro      = prop.getProperty("VAC_PF_MAX");
        }
        catch( Exception e){
            CZSystem.exit(-1,"CZTPGMain NO Propertie File");
        }

        ro_name = CZSystem.getRoName();
        ro_db_name = CZSystem.getDBName();

        setTitle("�s�o�f");                         //���Title
        setSize(1152,920);                          //��ʃT�C�Y
        setResizable(false);                        //��ʂ̃T�C�Y�ύX�͕s��
        setModal(true);                             //Modal�ŕ\��
        getContentPane().setLayout(null);
        // ����n�Q�Ƌ@�\    @20131021
        if(CZSystemDefine.REFERENCE_RUN == CZSystem.getRunLevel()){
            getContentPane().setBackground(CZSystemDefine.DEFAULT_REFERENCE_BACKGROUND_COL);
        }else{
            getContentPane().setBackground(CZSystemDefine.DEFAULT_BACKGROUND_COL);
        }

        //�F�ԕ\��
		String s = CZSystem.RoKetaChg(ro_name);	// 20050725 �F�F�\�������ύX
        main_ro_name_lab = new JLabel(s,JLabel.CENTER);
//        main_ro_name_lab = new JLabel(ro_name,JLabel.CENTER);
        main_ro_name_lab.setBounds(20, 20, 100, 30);
        main_ro_name_lab.setLocale(new Locale("ja","JP"));
        main_ro_name_lab.setFont(new java.awt.Font("dialog", 0, 18));
        main_ro_name_lab.setBorder(new Flush3DBorder());
        main_ro_name_lab.setForeground(java.awt.Color.black);
        getContentPane().add(main_ro_name_lab);

        //�����p�l��
        conpane = new ConditionPanel();
        conpane.setBounds(20, 60, 100, 340);
        getContentPane().add(conpane);

        //�O���t�ݒ�p�l��
        grapane = new GraphPanel();
        grapane.setBounds(20, 410, 100, 110);
        getContentPane().add(grapane);

        //�}�E�X���W�\���p�l��
        main_mouse_view = new MainMouseView();
        main_mouse_view.setBounds(20, 570, 100, 300);
        getContentPane().add(main_mouse_view);

        //�ȈՃO���t�\���p�l��
        simpgrapane = new SimplGraphPanel(1000,300);
        simpgrapane.setBounds(140, 570, 1000, 300);
        getContentPane().add(simpgrapane);

        // �O���t�\���̈�
        PVGrEventCompo comp = new PVGrEventCompo();

        main_sc = new MainSc(comp);                 // ���C���O���t�̃p�l��
        main_sc.setBounds(190, 20, 890, 500);
        main_sc.setDefault();
        getContentPane().add(main_sc);

        x_sc    = new XSc(comp);                    // X���̖ڐ��̃p�l��
        x_sc.setBounds(190, 520, 890, 40);
        x_sc.setDefault();
        getContentPane().add(x_sc);

        y1_sc   = new Y1Sc(comp);                   // Y���̍����̃p�l��
        y1_sc.setBounds(140, 20, 50, 500);
        y1_sc.setDefault();
        getContentPane().add(y1_sc);

        y2_sc   = new Y2Sc(comp);                   // Y���̉E���̃p�l��
        y2_sc.setBounds(1080, 20, 60, 500);
        y2_sc.setDefault();
        getContentPane().add(y2_sc);

        serche_dia = new SercheDialog();            //����Dialog
        serche_dia.setVisible(false);

        y_leng_dia = new YLengDialog();             //���ڐݒ�Dialog
        y_leng_dia.setVisible(false);

        CZSystem.log("CZTPGMain","CZTPGMain new");
    }

    // �F�Ԃ�DB���̂��擾����B
    // @return true ... OK, false ... NG
    public boolean setDefault(){
        ro_name = CZSystem.getRoName();
        ro_db_name = CZSystem.getDBName();
        return true;
    }

    // �F�ԕ\����\������B
    //
    private void setMainRoName(){

		String s = CZSystem.RoKetaChg(ro_name);	// 20050725 �F�F�\�������ύX
        main_ro_name_lab.setText(s);
//        main_ro_name_lab.setText(ro_name);
    }

    // X���̒�����ύX����B
    // @param len ����
    private void chgXLength(String len){
        gr_x_length = len;
        main_sc.chgXSize();
        x_sc.chgXSize();
        simpgrapane.chgXSize();
    }

    // Y���̒�����ύX����B
    //
    private void chgYLength(){
        main_sc.chgYSize();
        y1_sc.chgYSize();
        y2_sc.chgYSize();
        simpgrapane.chgYLength();
    }

    // ���\����ύX����B
    //
    private void chgShld(){
        main_sc.chgYSize();
        simpgrapane.chgShld();
    }

    // �s�o�f�G���[���b�Z�[�W�\��Dialog
    // @param msg ... ���b�Z�[�W���e
    // @return true ... OK, false ... NG
    private boolean errorMsg(Object msg[]){
        JOptionPane.showMessageDialog(null,msg,
                                "�s�o�f�G���[",
                                JOptionPane.ERROR_MESSAGE);
        return true;
    }

    // �o�b�`�J�n������ݒ肷��
    // @param st ... �o�b�`�J�n����
    // @return true ... OK, false ... NG
    public boolean setBtStart(CZSystemStart st){
        ro_bt_start = st;
        if(null == ro_bt_start) return false;
        return true;
    }

    // �ݒ�ς݃o�b�`�J�n�������폜����
    // @return true ... OK
    public boolean removeBtStart(){
        ro_bt_start = null;
        return true;
    }

    //
    // PV�f�[�^��ǂݍ���
    // @return ... �{�f�B�[���т̓Ǎ��݌���
    // �i-1 ... ���і���,-2 ... �\����,-3 ... �V�����_�[���і���,-4 ... �{�f�B�[���і����j
    public int readBtPV(){
        if(null == ro_bt_start){
            Object msg[] = {"�X�^�[�g���т��L��܂���I�I",
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
            Object msg[] = {"�\�����݂��܂���I�I",
                            view,
                            ""};
            errorMsg(msg);
            return -2;
        }

        boolean data_no[] = new boolean[CZSystemDefine.PV_MAX_LENGTH];
        for(int i = 0 ; i < CZSystemDefine.PV_MAX_LENGTH ; i++) data_no[i] = false;

        data_no[ MAIN1_H_T  ]   = true;   // 15   ���C���q�[�^�[�P���x
        data_no[ MAIN1_H_T_PF ] = true;   // 67   ���C���q�[�^�[�P���x�v���t�@�C��
        data_no[ DIA        ]   = true;   // 25   ���a
        data_no[ DIA_PF     ]   = true;   // 24   ���a�v���t�@�C��
        data_no[ SXL_ST     ]   = true;   // 18   �����グ���x
        data_no[ SXL_ST_PF  ]   = true;   // 76   �����グ���x�v���t�@�C��

        data_no[ SXL_RT     ]   = true;   // 19   �V�[�h��]
        data_no[ SXL_RT_PF  ]   = true;   // 81   �V�[�h��]�v���t�@�C��
        data_no[ CRU_RT     ]   = true;   // 21   ���c�{��]
        data_no[ CRU_RT_PF  ]   = true;   // 87   ���c�{��]�v���t�@�C��

        data_no[ PULL_AR    ]   = true;   // 16   �v���A���S��
        data_no[ PULL_AR_PF ]   = true;   // 72   �v���A���S���v���t�@�C��
        data_no[ VAC        ]   = true;   // 33   �F����
        data_no[ VAC_PF     ]   = true;   // 89   �F�����v���t�@�C��

        //�V�����_�[�ǂݍ���
        pv_data_shld = CZSystem.getPVData(ro_db_name,view,st.p_renban-1,data_no);
//@@        CZSystem.log("CZTPGMain readBtPV ","pv_data_shld  [" + pv_data_shld.size()  + "]");
        if(1 > pv_data_shld.size()){
            Object msg[] = {"�V�����_�[���т��L��܂���I�I",
                            "[" + pv_data_shld.size() + "]",
                            ""};
            errorMsg(msg);
            pv_data_shld = null;
            return -3;
        }

        //�{�f�B�[�ǂݍ���
        pv_data_body = CZSystem.getPVData(ro_db_name,view,st.p_renban,data_no);
//@@        CZSystem.log("CZTPGMain readBtPV ","pv_data_body  [" + pv_data_body.size()  + "]");
        if(1 > pv_data_body.size()){
            Object msg[] = {"�{�f�B�[���т��L��܂���I�I",
                            "[" + pv_data_body.size() + "]",
                            ""};
            errorMsg(msg);
            pv_data_body = null;
            return -4;
        }
        return pv_data_body.size();
    }

    // �J�[�\����ݒ肷��B
    //
    private void setCur(Cursor cu){
        serche_dia.setCursor(cu);
    }

    //
    // �J�[�\�����擾����B
    private Cursor getCur(){
        return serche_dia.getCursor();
    }

    /*******************************************************
     *
     *   �����̃p�l��
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

        // ---------- �R���X�g���N�^ -----------------------
        //
        ConditionPanel(){
            super();
            setName("ConditionPanel");
            setLayout(null);
            setBorder(new Flush3DBorder());
            // ����n�Q�Ƌ@�\    @20131021
            if(CZSystemDefine.REFERENCE_RUN == CZSystem.getRunLevel()){
                setBackground(CZSystemDefine.DEFAULT_REFERENCE_BACKGROUND_COL);
            }else{
                setBackground(CZSystemDefine.DEFAULT_BACKGROUND_COL);
            }

            int x = 10;
            int y = 10;
            int inc = 0;

            JButton search_button = new JButton("��  ��");
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

//@@�ǉ���
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
//@@�ǉ���

//@@            y += 40 ;
            y += inc;                   //@@
            JButton btcondition_button = new JButton("�������");
            btcondition_button.setBounds(x, y, 80, 24);
            btcondition_button.setLocale(new Locale("ja","JP"));
            btcondition_button.setFont(new java.awt.Font("dialog", 0, 18));
            btcondition_button.setBorder(new Flush3DBorder());
            btcondition_button.setForeground(java.awt.Color.black);
            btcondition_button.addActionListener(new BtConditionButton());
            add(btcondition_button);

//@@            y += 50 ;
            y += inc;                   //@@
            JButton controltable_button = new JButton("����e�[�u��");
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
        // T1�`T6�̐ݒ�
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
                t1_text.setText(String.valueOf(bt.no_youkai));      // �n��
                t2_text.setText(String.valueOf(bt.no_hikiage));     // ����
                t3_text.setText(String.valueOf(bt.no_kaiten));      // ��]
                t4_text.setText(String.valueOf(bt.no_toridasi));    // ��o
                t5_text.setText(String.valueOf(bt.no_aturyoku));    // ����
                t6_text.setText(String.valueOf(bt.no_teisu));       // �萔 @@�ǉ�
            }
            else{
                bt_text.setText("");
                t1_text.setText("");
                t2_text.setText("");
                t3_text.setText("");
                t4_text.setText("");
                t5_text.setText("");
                t6_text.setText("");        //@@�ǉ�
            }
        }

        /***************************************************
         *   �����{�^���̏���
         *    ����Dialog��\������B
         ***************************************************/
        class SearchButton implements ActionListener {
            public void actionPerformed(ActionEvent ev){
                serche_dia.setDefault();
                serche_dia.setVisible(true);
//@@                CZSystem.log("ConditionPanel SaveButton","actionPerformed");
            }
        }

        /***************************************************
         *   �����グ�����{�^��
         *    �����グ�����ݒ�Dialog��\������B
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
         *   ����e�[�u���ݒ�{�^��
         *    ����e�[�u���ݒ�Dialog��\������B
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
         * �����グ����Dialog
         *
         ***************************************************/
        class BtConditionDialog extends JDialog {

            // ---------- �R���X�g���N�^ -------------------
            //
            BtConditionDialog(){
                super();

                setTitle("�����グ����");
                setSize(820,250);
                setResizable(false);
                setModal(true);
                getContentPane().setLayout(null);
                // ����n�Q�Ƌ@�\    @20131021
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
             *       �a���o�^���ꗗ
             * @@T6�ǉ�
             ***********************************************/
            class BtConditionTable extends JTable {

                private Vector  bt_list     = null;

                private BtConditionTblMdl model = null;

                // ---------- �R���X�g���N�^ ---------------
                // @param v ... �o�b�`���
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
                 *       �a���o�^���ꗗ�F���f��
                 *
                 *******************************************/
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

                    // ---------- �R���X�g���N�^ -----------
                    // @param v ... �o�b�`���
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
                            data[i][14] = new Integer(bt.no_teisu);     //T6 @@
                            data[i][15] = new Integer(bt.pno_start);    //PNo
                            data[i][16] = new Integer(bt.p_kaisi);      //�J�n
                        }
                    }

                    // �������擾����B
                    // @return ... ����
                    public int getColumnCount(){
                        return TBL_COL;
                    }

                    // �s�����擾����B
                    // @return ... �s��
                    public int getRowCount(){
                        return TBL_ROW;
                    }

                    // �f�[�^���擾����B
                    // @param ... row:�s, col:��
                    // @return ... �f�[�^
                    public Object getValueAt(int row, int col){
                        return data[row][col];
                    }

                    // �������擾����B
                    // @param ... column:��
                    // @return ... ����
                    public String getColumnName(int column){
                        return names[column];
                    }

                    // �f�[�^�̌^���擾����B
                    // @param ... c:��
                    // @return ... �f�[�^�̌^
                    public Class getColumnClass(int c){
                        return getValueAt(0, c).getClass();
                    }

                    // cell�ҏW�̉ۂ��擾����B
                    // @param ... row:�s, col:��
                    // @return ... ����
                    public boolean isCellEditable(int row, int col){
                        return false;
                    }

                    // �f�[�^��ݒ肷��B
                    // @param ... aValue:�f�[�^, row:�s, col:��
                    // @return ... ����
                    public void setValueAt(Object aValue, int row, int column){
                        data[row][column] = aValue;
                    }
                } // BtConditionTblMdl
            } // BtConditionTable
        } // BtConditionDialog
    } // ConditionPanel


    /*******************************************************
     *
     *   �x���ݒ�A�w���ݒ�A�V�����_�[�̐ݒ�p�l��
     *
     *******************************************************/
    public class GraphPanel extends JPanel {

        JCheckBox shld_chk = null;

        // ---------- �R���X�g���N�^ -----------------------
        GraphPanel(){
            super();
            setName("GraphPanel");
            setLayout(null);
            setBorder(new Flush3DBorder());
            // ����n�Q�Ƌ@�\    @20131021
            if(CZSystemDefine.REFERENCE_RUN == CZSystem.getRunLevel()){
                setBackground(CZSystemDefine.DEFAULT_REFERENCE_BACKGROUND_COL);
            }else{
                setBackground(CZSystemDefine.DEFAULT_BACKGROUND_COL);
            }

            JButton leng_button = new JButton("���ڐݒ�");
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

            shld_chk = new JCheckBox("���\��",false);
            shld_chk.setBounds(10, 70, 80, 24);
            shld_chk.setLocale(new Locale("ja","JP"));
            shld_chk.setFont(new java.awt.Font("dialog", 0, 12));
            shld_chk.setBorderPaintedFlat(true);
            shld_chk.setForeground(java.awt.Color.black);
            shld_chk.addActionListener(new ShldChk());
            add(shld_chk);

//@@            CZSystem.log("GraphPanel GraphPanel","new");
        }

        // ���\���`�F�b�N�{�b�N�X�̃`�F�b�N
        // @return ... true, false
        public boolean isShld(){
            return shld_chk.isSelected();
        }

        /***************************************************
         * ���ڐݒ�{�^���̏���
         *  ���ڐݒ�Dialog��\������B
         ***************************************************/
        class ChgYLengButton implements ActionListener {
            public void actionPerformed(ActionEvent ev){
                y_leng_dia.setDefault();
                y_leng_dia.setVisible(true);
//@@                CZSystem.log("GraphPanel ChgYLengButton","actionPerformed");
            }
        } //ChgYLengButton

        /***************************************************
         * X���ڐ��ݒ�R���{�{�b�N�X�̏���
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
         * ���\���`�F�b�N�{�b�N�X�`�F�b�N���̏���
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
     *   ��]�A���͌n�ȈՃO���t
     *
     *******************************************************/
    public class SimplGraphPanel extends JPanel {
        private RotationPanel   r_view  = null;     // ��]
        private PressurePanel   p_view  = null;     // ����

        // ---------- �R���X�g���N�^ -----------------------
        // @param w ... ��, h ... ����
        SimplGraphPanel(int w,int h){
            super();
            setName("SimplGraphPanel");
            setLayout(null);
            setBorder(new Flush3DBorder());
            // ����n�Q�Ƌ@�\    @20131021
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

            JLabel lab = new JLabel("��]�n",JLabel.CENTER);
            lab.setBounds(x, 0, 50, y+2);
            lab.setLocale(new Locale("ja","JP"));
            lab.setFont(new java.awt.Font("dialog", 0, 12));
            lab.setForeground(java.awt.Color.black);
            add(lab);

            x = (w / 2) + (x / 2);
            p_view = new PressurePanel();
            p_view.setBounds(x, y, width, height);
            add(p_view);

            lab = new JLabel("���͌n",JLabel.CENTER);
            lab.setBounds(x, 0, 50, y+2);
            lab.setLocale(new Locale("ja","JP"));
            lab.setFont(new java.awt.Font("dialog", 0, 12));
            lab.setForeground(java.awt.Color.black);
            add(lab);

//@@            CZSystem.log("SimplGraphPanel SimplGraphPanel","new");
        }

        // �f�[�^��ݒ肵�A�ĕ\������B
        //
        public void setData(){
            r_view.setData();
            r_view.repaint();
            p_view.setData();
            p_view.repaint();
        }

        // X����ύX����B
        //
        public void chgXSize(){
            r_view.setData();
            r_view.repaint();
            p_view.setData();
            p_view.repaint();
        }

        // Y����ύX����B
        //
        private void chgYLength(){
            r_view.setData();
            r_view.repaint();
            p_view.setData();
            p_view.repaint();
        }

        // ���\����ύX����B
        //
        private void chgShld(){
            r_view.setData();
            r_view.repaint();
            p_view.setData();
            p_view.repaint();
        }

        /***************************************************
         * ��]��\������p�l��
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

            // ---------- �R���X�g���N�^ -------------------
            RotationPanel(){
                super();
                setName("SimplGraphPanel");
                setLayout(null);
                setBackground(BACK_COL);
            }

            //
            // �O���t��`�悷��B
            //
            public void paint(Graphics g){
                Dimension d = getSize(null);

                g.setColor(BACK_COL);
                g.fillRect(0,0,d.width,d.height);

                drawMemLine(g);     // �ڐ���
                drawLine(g);        // �O���t��
            }

            //
            // PV�l�����W���v�Z����B
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

                //�w�����W�v�Z�i���j
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

                //�w�����W�v�Z
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

                //�V�[�h��]�v���t�@�C���i���j
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

                //�V�[�h��]�v���t�@�C��
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

                //�V�[�h��]�i���j
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

                //�V�[�h��]
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

                //���c�{��]�v���t�@�C���i���j
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

                //���c�{��]�v���t�@�C��
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

                //���c�{��]�i���j
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

                //���c�{��]
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

            // X���̍��W���v�Z����B
            // @param w .. ��,min .. �ŏ��l,max .. �ő�l,val ..�f�[�^
            // @return X���̍��W
            private float xPos(int w,float min,float max,float val){
                float x_dot = (w - offset_x) / (max - min);
                float x = x_dot * (val - min) + offset_x;
                return x;
            }

            // Y���̍��W���v�Z����B
            // @param h .. ����,min .. �ŏ��l,max .. �ő�l,val ..�f�[�^
            // @return Y���̍��W
            private float yPos(int h,float min,float max,float val){
                float y_dot = (h - offset_y) / (max - min);
                float y = h - y_dot * (val - min) - offset_y;
                return y;
            }

            //
            // �O���t��������
            //
            private void drawLine(Graphics g){
                if(null == pv_data_shld) return;
                int size_shld = pv_data_shld.size();
                if(2 > size_shld) return;

                if(null == pv_data_body) return;
                int size = pv_data_body.size();
                if(2 > size) return;

                //�V�[�h��]�v���t�@�C���i���j
                if(grapane.isShld()){
                    g.setColor(java.awt.Color.white);
                    g.drawPolyline(x_pos_shld,y_pos_shld_sxl_rt_pf,size_shld);
                }
                //�V�[�h��]�v���t�@�C��
                g.setColor(java.awt.Color.white);
                g.drawPolyline(x_pos,y_pos_sxl_rt_pf,size);
                //�V�[�h��]�i���j
                if(grapane.isShld()){
                    g.setColor(SXL_RT_COL);
                    g.drawPolyline(x_pos_shld,y_pos_shld_sxl_rt,size_shld);
                }
                //�V�[�h��]
                g.setColor(SXL_RT_COL);
                g.drawPolyline(x_pos,y_pos_sxl_rt,size);
                g.drawString("SXL.RT",x_pos[size-1],y_pos_sxl_rt[size-1]);

                //���c�{��]�v���t�@�C���i���j
                if(grapane.isShld()){
                    g.setColor(java.awt.Color.white);
                    g.drawPolyline(x_pos_shld,y_pos_shld_cru_rt_pf,size_shld);
                }
                //���c�{��]�v���t�@�C��
                g.setColor(java.awt.Color.white);
                g.drawPolyline(x_pos,y_pos_cru_rt_pf,size);
                //���c�{��]�i���j
                if(grapane.isShld()){
                    g.setColor(CRU_RT_COL);
                    g.drawPolyline(x_pos_shld,y_pos_shld_cru_rt,size_shld);
                }
                //���c�{��]
                g.setColor(CRU_RT_COL);
                g.drawPolyline(x_pos,y_pos_cru_rt,size);
                g.drawString("CRU.RT",x_pos[size-1],y_pos_cru_rt[size-1]);
            }

            //
            // �ڐ���������
            //
            private void drawMemLine(Graphics g){
                float x;
                float y;
                float inc;

                Dimension d = getSize(null);

                // �w���ڐ� ��
                g.setColor(MEM_LINE3_COL);
                inc = (d.width - offset_x) / (gr_x_bun * 4.0f);
                for(x = offset_x ; x < d.width ; x+=inc){
                    g.drawLine((int)x,0,(int)x,d.height-offset_y);
                }
                // �x���ڐ�
                inc = (d.height - offset_y) / (gr_y_bun * 4.0f) ;
                for(y = d.height - offset_y ; y > 0 ; y-=inc){
                    g.drawLine((int)offset_x,(int)y,d.width,(int)y);
                }

                // �w���ڐ� ��
                g.setColor(MEM_LINE2_COL);
                inc = (d.width - offset_x) / (gr_x_bun * 2.0f);
                for(x = offset_x ; x < d.width ; x+=inc){
                    g.drawLine((int)x,0,(int)x,d.height-offset_y);
                }
                // �x���ڐ�
                inc = (d.height - offset_y) / (gr_y_bun * 2.0f) ;
                for(y = d.height - offset_y ; y > 0 ; y-=inc){
                    g.drawLine((int)offset_x,(int)y,d.width,(int)y);
                }

                // �w���ڐ� ��
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

                // �x���ڐ�
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
         * ���̓p�l��
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

            // ---------- �R���X�g���N�^ -------------------
            PressurePanel(){
                super();
                setName("SimplGraphPanel");
                setLayout(null);
                setBackground(BACK_COL);
            }

            //
            // �O���t��`�悷��B
            //
            public void paint(Graphics g){
                Dimension d = getSize(null);

                g.setColor(BACK_COL);
                g.fillRect(0,0,d.width,d.height);
                drawMemLine(g);
                drawLine(g);
            }

            //
            // �f�[�^������W���v�Z����B
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

                //�w�����W�v�Z�i���j
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

                //�w�����W�v�Z
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

                //�v���A���S���v���t�@�C���i���j
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

                //�v���A���S���v���t�@�C��
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

                //�v���A���S���i���j
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

                //�v���A���S��
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

                //�F�����v���t�@�C���i���j
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

                //�F�����v���t�@�C��
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

                //�F�����i���j
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

                //�F����
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

            // X���̍��W���v�Z����B
            // @param w .. ��,min .. �ŏ��l,max .. �ő�l,val ..�f�[�^
            // @return X���̍��W
            private float xPos(int w,float min,float max,float val){
                float x_dot = (w - offset_x) / (max - min);
                float x = x_dot * (val - min) + offset_x;
                return x;
            }

            // Y���̍��W���v�Z����B
            // @param h .. ��,min .. �ŏ��l,max .. �ő�l,val ..�f�[�^
            // @return Y���̍��W
            private float yPos(int h,float min,float max,float val){
                float y_dot = (h - offset_y) / (max - min);
                float y = h - y_dot * (val - min) - offset_y;
                return y;
            }

            //
            // �O���t��������
            //
            private void drawLine(Graphics g){
                if(null == pv_data_shld) return;
                int size_shld = pv_data_shld.size();
                if(2 > size_shld) return;

                if(null == pv_data_body) return;
                int size = pv_data_body.size();
                if(2 > size) return;

                //�v���A���S���v���t�@�C���i���j
                if(grapane.isShld()){
                    g.setColor(java.awt.Color.white);
                    g.drawPolyline(x_pos_shld,y_pos_shld_pull_ar_pf,size_shld);
                }
                //�v���A���S���v���t�@�C��
                g.setColor(java.awt.Color.white);
                g.drawPolyline(x_pos,y_pos_pull_ar_pf,size);

                //�v���A���S���i���j
                if(grapane.isShld()){
                    g.setColor(PULL_AR_COL);
                    g.drawPolyline(x_pos_shld,y_pos_shld_pull_ar,size_shld);
                }
                //�v���A���S��
                g.setColor(PULL_AR_COL);
                g.drawPolyline(x_pos,y_pos_pull_ar,size);
                g.drawString("PULL.AR",x_pos[size-1],y_pos_pull_ar[size-1]);

                //�F�����v���t�@�C���i���j
                if(grapane.isShld()){
                    g.setColor(java.awt.Color.white);
                    g.drawPolyline(x_pos_shld,y_pos_shld_vac_pf,size_shld);
                }
                //�F�����v���t�@�C��
                g.setColor(java.awt.Color.white);
                g.drawPolyline(x_pos,y_pos_vac_pf,size);

                //�F�����i���j
                if(grapane.isShld()){
                    g.setColor(VAC_COL);
                    g.drawPolyline(x_pos_shld,y_pos_shld_vac,size_shld);
                }
                //�F����
                g.setColor(VAC_COL);
                g.drawPolyline(x_pos,y_pos_vac,size);
                g.drawString("VAC",x_pos[size-1],y_pos_vac[size-1]);
            }

            //
            // �ڐ���������
            //
            private void drawMemLine(Graphics g){
                float x;
                float y;
                float inc;

                Dimension d = getSize(null);

                // �w���ڐ� ��
                g.setColor(MEM_LINE3_COL);
                inc = (d.width - offset_x) / (gr_x_bun * 4.0f);
                for(x = offset_x ; x < d.width ; x+=inc){
                    g.drawLine((int)x,0,(int)x,d.height-offset_y);
                }
                // �x���ڐ�
                inc = (d.height - offset_y) / (gr_y_bun * 4.0f) ;
                for(y = d.height - offset_y ; y > 0 ; y-=inc){
                    g.drawLine((int)offset_x,(int)y,d.width,(int)y);
                }

                // �w���ڐ� ��
                g.setColor(MEM_LINE2_COL);
                inc = (d.width - offset_x) / (gr_x_bun * 2.0f);
                for(x = offset_x ; x < d.width ; x+=inc){
                    g.drawLine((int)x,0,(int)x,d.height-offset_y);
                }
                // �x���ڐ�
                inc = (d.height - offset_y) / (gr_y_bun * 2.0f) ;
                for(y = d.height - offset_y ; y > 0 ; y-=inc){
                    g.drawLine((int)offset_x,(int)y,d.width,(int)y);
                }

                // �w���ڐ� ��
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

                // �x���ڐ�
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
     *   �}�E�X���W�\���p�l��
     *
     *******************************************************/
    public class MainMouseView extends JScrollPane {

        SubView view = null;    

        // ----------- �R���X�g���N�^ ----------------------
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
         * �f�[�^�\���p�l��
         ***************************************************/
        class SubView extends JPanel {

            // ---------- �R���X�g���N�^ -------------------
            SubView(){
                super();
                setLayout(null);
                setBackground(BACK_COL);
//@@                CZSystem.log("MainMouseView SubView","new");
            }

            //
            // �}�E�X�ʒu�̃f�[�^��\������B
            //
            public void paint(Graphics g){
                Dimension d = getSize(null);

                g.setColor(BACK_COL);
                g.fillRect(0,0,d.width,d.height);

                int x = 15;
                int y = 30;
                int inc = 20;

                //�w��
                g.setColor(MEM_LINE1_COL);
                g.drawString(String.valueOf(x_length_mouse),x,y);
                y+=inc;

                //���C���q�[�^�[�P���x
                y+=inc;
                g.setColor(MAIN1_H_T_COL);
                g.drawString(String.valueOf(y_main1_h_t_mouse),x,y);

                //���C���q�[�^�[�P���x�v���t�@�C��
                y+=inc;
                g.setColor(MAIN1_H_T_PF_COL);
                g.drawString(String.valueOf(y_main1_h_t_pf_mouse),x,y);

                //���a
                y+=inc;
                g.setColor(DIA_COL);
                g.drawString(String.valueOf(y_dia_mouse),x,y);

                //�����グ���x
                y+=inc;
                g.setColor(SXL_ST_COL);
                g.drawString(String.valueOf(y_sxl_st_mouse),x,y);

                //�����グ���x�v���t�@�C��
                y+=inc;
                g.setColor(SXL_ST_PF_COL);
                g.drawString(String.valueOf(y_sxl_st_pf_mouse),x,y);
            }
        } // SubView
    } // MainMouseView


    /*******************************************************
     *
     * ����Dialog
     *
     *******************************************************/
    class SercheDialog extends JDialog {

        private JScrollPane bt_scpanel      = null;
        private JScrollPane bt_start_scpanel    = null;
        private JButton     read_button     = null;
        private JLabel      ro_name_lab     = null;

        //
        // ---------- �R���X�g���N�^ -----------------------
        //
        SercheDialog(){
            super();

            setTitle("��  ��");
            setSize(820,335);
            setResizable(false);
            setModal(true);
            getContentPane().setLayout(null);
            // ����n�Q�Ƌ@�\    @20131021
            if(CZSystemDefine.REFERENCE_RUN == CZSystem.getRunLevel()){
                getContentPane().setBackground(CZSystemDefine.DEFAULT_REFERENCE_BACKGROUND_COL);
            }else{
                getContentPane().setBackground(CZSystemDefine.DEFAULT_BACKGROUND_COL);
            }

			String s = CZSystem.RoKetaChg(ro_name);	// 20050725 �F�F�\�������ύX
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

            read_button = new JButton("�ǂݍ���");
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
        // �o�b�`����\������B
        // @return true
        public boolean setDefault(){
            removeBtStart();
            removeBtCondition();
			String s = CZSystem.RoKetaChg(ro_name);	// 20050725 �F�F�\�������ύX
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
        // �o�b�`����ݒ肷��B
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
        // �o�b�`�����폜����B
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
         * �Ǎ��݃{�^���̏���
         *  �o�b�`�����ēǍ�����B
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
         * �o�b�`���̈ꗗ��\������B
         *
         ***************************************************/
        class BtTable extends JTable {

            private Vector  bt_all_list     = null;
            private Vector  bt_list         = null;

            private BtTblMdl model = null;

            private boolean life = false;

            // ---------- �R���X�g���N�^ -------------------
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

                    // �o�^����
                    colum = cmdl.getColumn(2);
                    colum.setMaxWidth(162);
                    colum.setMinWidth(162);
                    colum.setWidth(162);

                }
                catch (Throwable e) {
                    CZSystem.handleException(e);
                }
            }

            // �o�b�`���I�����̏���
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
         * �o�b�`�����шꗗ�F���f��
         *
         ***************************************************/
        public class BtTblMdl extends AbstractTableModel {

            private int TBL_ROW             = 0;
            final   int TBL_COL             = 3;
            private Vector  bt_list         = null;

            final String[] names = {" # "  , "Bt" , "�o�^����" };

            private Object  data[][];

            // ---------- �R���X�g���N�^ -------------------
            // @param v �o�b�`���
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
                    data[i][2] = bt.t_time;         // �o�^����
                }
            }

            // �񐔂��擾����B
            // @return ��
            public int getColumnCount(){
                return TBL_COL;
            }

            // �s�����擾����B
            // @return �s��
            public int getRowCount(){
                return TBL_ROW;
            }

            // �l���擾����B
            // @param row ... �s, col ... ��
            // @return �l
            public Object getValueAt(int row, int col){
                return data[row][col];
            }

            // �񖼂��擾����B
            // @param column ... ��
            // @return ��
            public String getColumnName(int column){
                return names[column];
            }

            // ��̃f�[�^�^���擾����B
            // @param c ... ��
            // @return �f�[�^�̌^
            public Class getColumnClass(int c){
                return getValueAt(0, c).getClass();
            }

            // �Z���̕ҏW�ۂ��擾����B
            // @param row ... �s, col ... ��
            // @return true :��, false:��
            public boolean isCellEditable(int row, int col){
                return false;
            }

            // �l��ݒ肷��B
            // @param aValue ... �l, row ... �s, column ... ��
            public void setValueAt(Object aValue, int row, int column){
                data[row][column] = aValue;
            }
        } // BtTblMdl


        /***************************************************
         *
         *       �a���X�^�[�g���Ԉꗗ
         *
         ***************************************************/
        class BtStartTable extends JTable {

            private Vector  bt_list         = null;
            private Vector  bt_start_list   = null;

            private BtStartTblMdl model = null;

            private boolean life = false;

            // ---------- �R���X�g���N�^ -------------------
            // @param v �o�b�`���
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

                    //NULL���K�v
                    if(null == tmp) return;

                    //Body �����ɂ���
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

            //
            // �I�����̏���
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
         *       �a���X�^�[�g���Ԉꗗ�F���f��
         *
         ***************************************************/
        public class BtStartTblMdl extends AbstractTableModel {

            private int TBL_ROW             = 0;
            final   int TBL_COL             = 6;
            private Vector  bt_start_list   = null;

            final String[] names = {" # "  , "PNo" ,
                                        "SPNo","PSeq"  ,
                                        "�v���Z�X",
                                        "�o�^����" };
            private Object  data[][];

            // ---------- �R���X�g���N�^ -------------------
            // @param v �o�b�`���
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
                    data[i][4] = CZSystem.getProcName(st.p_no);     // �v���Z�X
                    data[i][5] = st.p_start;                        // �o�^����
                }
            }

            // �񐔂��擾����B
            // @return ��
            public int getColumnCount(){
                return TBL_COL;
            }

            // �s�����擾����B
            // @return �s��
            public int getRowCount(){
                return TBL_ROW;
            }

            // �l���擾����B
            // @param row .. �s, col .. ��
            // @return �l
            public Object getValueAt(int row, int col){
                return data[row][col];
            }

            // �񖼂��擾����B
            // @param column ... ��
            // @return ��
            public String getColumnName(int column){
                return names[column];
            }

            // �f�[�^�̌^���擾����B
            // @param c ... ��
            // @return �f�[�^�̌^
            public Class getColumnClass(int c){
                return getValueAt(0, c).getClass();
            }

            // �ҏW�ۂ��擾����B
            // @param row .. �s,col .. ��
            // @return true .. ��, false .. ��
            public boolean isCellEditable(int row, int col){
                return false;
            }

            // �񐔂��擾����B
            // @param aValue .. ,row .. �s,column .. ��
            public void setValueAt(Object aValue, int row, int column){
                data[row][column] = aValue;
            }
        }
    } // SercheDialog

    /*******************************************************
     *  Y�����ڐݒ�Dialog
     *  @@ T6��ǉ�����K�v����
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

        private JLabel yLen_ro_name_lab = null;     //�F�ԕ\��

        //
        // ---------- �R���X�g���N�^ -----------------------
        YLengDialog(){
            super();

            setTitle("���ڐݒ�");
            setSize(440,565);
            setResizable(false);
            setModal(true);
            getContentPane().setLayout(null);
            // ����n�Q�Ƌ@�\    @20131021
            if(CZSystemDefine.REFERENCE_RUN == CZSystem.getRunLevel()){
                getContentPane().setBackground(CZSystemDefine.DEFAULT_REFERENCE_BACKGROUND_COL);
            }else{
                getContentPane().setBackground(CZSystemDefine.DEFAULT_BACKGROUND_COL);
            }

			String s = CZSystem.RoKetaChg(ro_name);	// 20050725 �F�F�\�������ύX
            yLen_ro_name_lab = new JLabel(s,JLabel.CENTER);
//            yLen_ro_name_lab = new JLabel(ro_name,JLabel.CENTER);
            yLen_ro_name_lab.setBounds(20, 20, 100, 30);
            yLen_ro_name_lab.setLocale(new Locale("ja","JP"));
            yLen_ro_name_lab.setFont(new java.awt.Font("dialog", 0, 18));
            yLen_ro_name_lab.setBorder(new Flush3DBorder());
            yLen_ro_name_lab.setForeground(java.awt.Color.black);
            getContentPane().add(yLen_ro_name_lab);

            JButton set_button = new JButton("��  ��");
            set_button.setBounds(320, 500, 100, 24);
            set_button.setLocale(new Locale("ja","JP"));
            set_button.setFont(new java.awt.Font("dialog", 0, 18));
            set_button.setBorder(new Flush3DBorder());
            set_button.setForeground(java.awt.Color.black);
            set_button.addActionListener(new SetButton());
            getContentPane().add(set_button);

            JLabel lab = null;

            lab = new JLabel("�l����",JLabel.CENTER);
            lab.setBounds(220, 70, 100, 24);
            lab.setLocale(new Locale("ja","JP"));
            lab.setFont(new java.awt.Font("dialog", 0, 16));
            lab.setBorder(new Flush3DBorder());
            lab.setForeground(java.awt.Color.black);
            getContentPane().add(lab);

            lab = new JLabel("�l����",JLabel.CENTER);
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
            lab = new JLabel("���C���q�[�^�[�P���x",JLabel.CENTER);
            lab.setBounds(20, y, 200, 24);
            lab.setLocale(new Locale("ja","JP"));
            lab.setFont(new java.awt.Font("dialog", 0, 16));
            lab.setBorder(new Flush3DBorder());
            lab.setForeground(java.awt.Color.black);
            getContentPane().add(lab);

            y+=inc;
            lab = new JLabel("���C���q�[�^�[�P���x�o�e",JLabel.CENTER);
            lab.setBounds(20, y, 200, 24);
            lab.setLocale(new Locale("ja","JP"));
            lab.setFont(new java.awt.Font("dialog", 0, 16));
            lab.setBorder(new Flush3DBorder());
            lab.setForeground(java.awt.Color.black);
            getContentPane().add(lab);

            y+=inc;
            lab = new JLabel("���a",JLabel.CENTER);
            lab.setBounds(20, y, 200, 24);
            lab.setLocale(new Locale("ja","JP"));
            lab.setFont(new java.awt.Font("dialog", 0, 16));
            lab.setBorder(new Flush3DBorder());
            lab.setForeground(java.awt.Color.black);
            getContentPane().add(lab);

            y+=inc;
            lab = new JLabel("�����グ���x",JLabel.CENTER);
            lab.setBounds(20, y, 200, 24);
            lab.setLocale(new Locale("ja","JP"));
            lab.setFont(new java.awt.Font("dialog", 0, 16));
            lab.setBorder(new Flush3DBorder());
            lab.setForeground(java.awt.Color.black);
            getContentPane().add(lab);

            y+=inc;
            lab = new JLabel("�����グ���x�o�e",JLabel.CENTER);
            lab.setBounds(20, y, 200, 24);
            lab.setLocale(new Locale("ja","JP"));
            lab.setFont(new java.awt.Font("dialog", 0, 16));
            lab.setBorder(new Flush3DBorder());
            lab.setForeground(java.awt.Color.black);
            getContentPane().add(lab);

            // �g��
            y+=inc;
            y+=inc;
            lab = new JLabel("���a�Ǘ�",JLabel.CENTER);
            lab.setBounds(20, y, 200, 24);
            lab.setLocale(new Locale("ja","JP"));
            lab.setFont(new java.awt.Font("dialog", 0, 16));
            lab.setBorder(new Flush3DBorder());
            lab.setForeground(java.awt.Color.black);
            getContentPane().add(lab);

            y+=inc;
            lab = new JLabel("�V�[�h��]",JLabel.CENTER);
            lab.setBounds(20, y, 200, 24);
            lab.setLocale(new Locale("ja","JP"));
            lab.setFont(new java.awt.Font("dialog", 0, 16));
            lab.setBorder(new Flush3DBorder());
            lab.setForeground(java.awt.Color.black);
            getContentPane().add(lab);

            y+=inc;
            lab = new JLabel("���c�{��]",JLabel.CENTER);
            lab.setBounds(20, y, 200, 24);
            lab.setLocale(new Locale("ja","JP"));
            lab.setFont(new java.awt.Font("dialog", 0, 16));
            lab.setBorder(new Flush3DBorder());
            lab.setForeground(java.awt.Color.black);
            getContentPane().add(lab);

            y+=inc;
            lab = new JLabel("�v���A���S��",JLabel.CENTER);
            lab.setBounds(20, y, 200, 24);
            lab.setLocale(new Locale("ja","JP"));
            lab.setFont(new java.awt.Font("dialog", 0, 16));
            lab.setBorder(new Flush3DBorder());
            lab.setForeground(java.awt.Color.black);
            getContentPane().add(lab);

            y+=inc;
            lab = new JLabel("�F����",JLabel.CENTER);
            lab.setBounds(20, y, 200, 24);
            lab.setLocale(new Locale("ja","JP"));
            lab.setFont(new java.awt.Font("dialog", 0, 16));
            lab.setBorder(new Flush3DBorder());
            lab.setForeground(java.awt.Color.black);
            getContentPane().add(lab);

            // ���͗̈�

            //
            // �q�[�^�[���x�֌W
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

            // ���a�x�֌W
            y+=inc;
            dia_min = new LengText();
            dia_min.setBounds(220, y, 100, 24);
            dia_min.setForeground(DIA_COL);
            getContentPane().add(dia_min);

            dia_max = new LengText();
            dia_max.setBounds(320, y, 100, 24);
            dia_max.setForeground(DIA_COL);
            getContentPane().add(dia_max);

            // �����グ���x�֌W
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

            // �g��
            // ���a�x�֌W
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

            // �V�[�h��]
            y+=inc;
            sxl_rt_min = new LengText();
            sxl_rt_min.setBounds(220, y, 100, 24);
            sxl_rt_min.setForeground(SXL_RT_COL);
            getContentPane().add(sxl_rt_min);

            sxl_rt_max = new LengText();
            sxl_rt_max.setBounds(320, y, 100, 24);
            sxl_rt_max.setForeground(SXL_RT_COL);
            getContentPane().add(sxl_rt_max);

            // ���c�{��]
            y+=inc;
            cru_rt_min = new LengText();
            cru_rt_min.setBounds(220, y, 100, 24);
            cru_rt_min.setForeground(CRU_RT_COL);
            getContentPane().add(cru_rt_min);

            cru_rt_max = new LengText();
            cru_rt_max.setBounds(320, y, 100, 24);
            cru_rt_max.setForeground(CRU_RT_COL);
            getContentPane().add(cru_rt_max);

            // �v���A���S��
            y+=inc;
            pull_ar_min = new LengText();
            pull_ar_min.setBounds(220, y, 100, 24);
            pull_ar_min.setForeground(PULL_AR_COL);
            getContentPane().add(pull_ar_min);

            pull_ar_max = new LengText();
            pull_ar_max.setBounds(320, y, 100, 24);
            pull_ar_max.setForeground(PULL_AR_COL);
            getContentPane().add(pull_ar_max);

            // �F����
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

        // ����l��ݒ肷��B
        //
        public boolean setDefault(){
			String s = CZSystem.RoKetaChg(ro_name);	// 20050725 �F�F�\�������ύX
            yLen_ro_name_lab.setText(s);
//            yLen_ro_name_lab.setText(ro_name);

            // �q�[�^�[���x�֌W
            ht_min.setText(main1_h_t_min_pro);
            ht_max.setText(main1_h_t_max_pro);
            ht_pf_min.setText(main1_h_t_pf_min_pro);
            ht_pf_max.setText(main1_h_t_pf_max_pro);

            // ���a�x�֌W
            dia_min.setText(dia_min_pro);
            dia_max.setText(dia_max_pro);
            dia_pf_min.setText(dia_pf_min_pro);
            dia_pf_max.setText(dia_pf_max_pro);

            // �����グ���x�֌W
            fp_min.setText(sxl_st_min_pro);
            fp_max.setText(sxl_st_max_pro);
            fp_pf_min.setText(sxl_st_pf_min_pro);
            fp_pf_max.setText(sxl_st_pf_max_pro);

            // �V�[�h��]
            sxl_rt_min.setText(sxl_rt_pf_min_pro);
            sxl_rt_max.setText(sxl_rt_pf_max_pro);

            // ���c�{��]
            cru_rt_min.setText(cru_rt_pf_min_pro);
            cru_rt_max.setText(cru_rt_pf_max_pro);

            // �v���A���S��
            pull_ar_min.setText(pull_ar_pf_min_pro);
            pull_ar_max.setText(pull_ar_pf_max_pro);

            // �F����
            vac_min.setText(vac_pf_min_pro);
            vac_max.setText(vac_pf_max_pro);
            return true;
        }

        // Y����ݒ肷��
        //
        private boolean setYLang(){

            // �q�[�^�[���x�֌W
            main1_h_t_min_pro       = ht_min.getText();
            main1_h_t_max_pro       = ht_max.getText();
            main1_h_t_pf_min_pro    = ht_pf_min.getText();
            main1_h_t_pf_max_pro    = ht_pf_max.getText();

            // ���a�x�֌W
            dia_min_pro = dia_min.getText();
            dia_max_pro = dia_max.getText();
            dia_pf_min_pro = dia_pf_min.getText();
            dia_pf_max_pro = dia_pf_max.getText();

            // �����グ���x�֌W
            sxl_st_min_pro = fp_min.getText();
            sxl_st_max_pro = fp_max.getText();
            sxl_st_pf_min_pro = fp_pf_min.getText();
            sxl_st_pf_max_pro = fp_pf_max.getText();

            // �V�[�h��]
            sxl_rt_pf_min_pro = sxl_rt_min.getText();
            sxl_rt_pf_max_pro = sxl_rt_max.getText();

            // ���c�{��]
            cru_rt_pf_min_pro = cru_rt_min.getText();
            cru_rt_pf_max_pro = cru_rt_max.getText();

            // �v���A���S��
            pull_ar_pf_min_pro = pull_ar_min.getText();
            pull_ar_pf_max_pro = pull_ar_max.getText();

            // �F����
            vac_pf_min_pro = vac_min.getText();
            vac_pf_max_pro = vac_max.getText();

            chgYLength();
            return true;
        }

        /***************************************************
         *
         *       �ݒ�{�^���̏���
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
         *       ����Min,Max����͂���TextField
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
     *       ���C���O���t
     *
     *******************************************************/
    public class MainSc extends JScrollPane {

        private Rectangle   view_rec    = null;
        private View        view        = null;

        // ---------- �R���X�g���N�^ -----------------------
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

        // X��
        //
        public void chgXSize(){
            Float x_def = new Float(GR_X_LENGTH_DEF);
            Float x_new = new Float(gr_x_length);
            float new_size = (x_def.floatValue()/x_new.floatValue()) * view_rec.width;
            
            Dimension d = view.getSize(null);
            view.setSize(new Dimension((int)new_size,d.height));
            setData();
        }

        // Y��
        //
        public void chgYSize(){
            setData();
        }

        //
        // �f�[�^��ݒ肵�A�O���t���ĕ`�悷��B
        //
        public void setData(){
            view.setData();
            view.repaint();
        }

        /***************************************************
         *  �O���t�`��p�l��
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

            // ---------- �R���X�g���N�^ -------------------
            //
            View(){
                super();
                setName("View");
                setLayout(null);
                setBackground(BACK_COL);
                addMouseMotionListener(new MainViewMouseMotion());
            }

            // �g��ݒ肷��B
            //
            public void setViewRec(Rectangle rec){
                view_rec = rec;
            }

            // �O���t��`�悷��B
            //
            public void paint(Graphics g){
                Dimension d = getSize(null);

                g.setColor(BACK_COL);
                g.fillRect(0,0,d.width,d.height);
                drawMemLine(g);
                drawLine(g);
            }

            // �ڐ�����`�悷��B
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

            // �O���t����`�悷��B
            //
            private void drawLine(Graphics g){
                if(null == pv_data_shld) return;
                int size_shld = pv_data_shld.size();
                if(2 > size_shld) return;

                if(null == pv_data_body) return;
                int size = pv_data_body.size();
                if(2 > size) return;


                //�����グ���x�i���j
                if(grapane.isShld()){
                    g.setColor(SXL_ST_COL);
                    g.drawPolyline(x_pos_shld,y_pos_shld_sxl_st,size_shld);
                }

                //�����グ���x
                g.setColor(SXL_ST_COL);
                g.drawPolyline(x_pos,y_pos_sxl_st,size);
                g.drawString("SXL.ST",x_pos[size-1],y_pos_sxl_st[size-1]);

                //�����グ���x�v���t�@�C���i���j
                if(grapane.isShld()){
                    g.setColor(SXL_ST_PF_COL);
                    g.drawPolyline(x_pos_shld,y_pos_shld_sxl_st_pf,size_shld);
                }

                //�����グ���x�v���t�@�C��
                g.setColor(SXL_ST_PF_COL);
                g.drawPolyline(x_pos,y_pos_sxl_st_pf,size);
                g.drawString("SXS.PF",x_pos[size-1],y_pos_sxl_st_pf[size-1]);

                //���C���q�[�^�[�P���x�i���j
                if(grapane.isShld()){
                    //���C���q�[�^�[�P���x�ƃv���t�@�C��
                    g.setColor(java.awt.Color.white);
                    g.drawPolyline(x_pos_shld,y_pos_shld_ht_conv,size_shld);
                    //���C���q�[�^�[�P���x
                    g.setColor(MAIN1_H_T_COL);
                    g.drawPolyline(x_pos_shld,y_pos_shld_ht,size_shld);
                }

                //���C���q�[�^�[�P���x
                //���C���q�[�^�[�P���x�ƃv���t�@�C��
                g.setColor(java.awt.Color.white);
                g.drawPolyline(x_pos,y_pos_ht_conv,size);
                //���C���q�[�^�[�P���x
                g.setColor(MAIN1_H_T_COL);
                g.drawPolyline(x_pos,y_pos_ht,size);
                g.drawString("HEA.T1",x_pos[size-1],y_pos_ht[size-1]);

                //���C���q�[�^�[�P���x�v���t�@�C���i���j
                if(grapane.isShld()){
                    //�v���t�@�C���ƃ��C���q�[�^�[�P
                    g.setColor(java.awt.Color.lightGray);
                    g.drawPolyline(x_pos_shld,y_pos_shld_ht_pf_conv,size_shld);
                    //���C���q�[�^�[�P���x�v���t�@�C��
                    g.setColor(MAIN1_H_T_PF_COL);
                    g.drawPolyline(x_pos_shld,y_pos_shld_ht_pf,size_shld);
                }

                //���C���q�[�^�[�P���x�v���t�@�C��
                //�v���t�@�C���ƃ��C���q�[�^�[�P
                g.setColor(java.awt.Color.lightGray);
                g.drawPolyline(x_pos,y_pos_ht_pf_conv,size);
                //���C���q�[�^�[�P���x�v���t�@�C��
                g.setColor(MAIN1_H_T_PF_COL);
                g.drawPolyline(x_pos,y_pos_ht_pf,size);
                g.drawString("HT1.PF",x_pos[size-1],y_pos_ht_pf[size-1]);

                //���a�v���t�@�C���i���j
                if(grapane.isShld()){
                    g.setColor(DIA_PF_COL);
                    g.drawPolyline(x_pos_shld,y_pos_shld_dia_pf_max,size_shld);
                    g.drawPolyline(x_pos_shld,y_pos_shld_dia_pf_min,size_shld);
                    g.drawPolyline(x_pos_shld,y_pos_shld_dia_pf,size_shld);
                }

                //���a�v���t�@�C��
                g.setColor(DIA_PF_COL);
                g.drawPolyline(x_pos,y_pos_dia_pf_max,size);
                g.drawPolyline(x_pos,y_pos_dia_pf_min,size);
                g.drawPolyline(x_pos,y_pos_dia_pf,size);
                g.drawString("DIA.PF",x_pos[size-1],y_pos_dia_pf[size-1]);

                //���a�i���j
                if(grapane.isShld()){
                    g.setColor(DIA_COL);
                    g.drawPolyline(x_pos_shld,y_pos_shld_dia,size_shld);
                }

                //���a
                g.setColor(DIA_COL);
                g.drawPolyline(x_pos,y_pos_dia,size);
                g.drawString("DIA",x_pos[size-1],y_pos_dia[size-1]);
            }

            // �f�[�^������W���v�Z����B
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

                //�w�����W�v�Z�i���j
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

                //�w�����W�v�Z
                x_pos = new int[size];
                x_max = new Float(gr_x_length);
                min = 0.0f;
                max = x_max.floatValue();

                for(int i = 0 ; i < size ; i++){
                    data = (CZSystemPVData)pv_data_body.elementAt(i);
                    val = data.p_length + val_shld;
                    x_pos[i] = (int)xPos(d.width,view_rec.width,min,max,val);
                }

                //���C���q�[�^�[�P���x�i���j
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

                //���C���q�[�^�[�P���x
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

                //���C���q�[�^�[�P���x�ƃv���t�@�C���i���j
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

                //���C���q�[�^�[�P���x�ƃv���t�@�C��
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

                //���C���q�[�^�[�P���x�v���t�@�C���i���j
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

                //���C���q�[�^�[�P���x�v���t�@�C��
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

                //���C���q�[�^�[�P���x�v���t�@�C���ƃ��C���q�[�^�[�P���x�i���j
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

                //���C���q�[�^�[�P���x�v���t�@�C���ƃ��C���q�[�^�[�P���x
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

                //���a�i���j
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

                //���a
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

                //���a�v���t�@�C���i���j
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

                //���a�v���t�@�C��
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

                //�����グ���x�i���j
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

                //�����グ���x
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

                //�����グ���x�v���t�@�C���i���j
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

                //�����グ���x�v���t�@�C��
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

            // �}�E�X���W�̃f�[�^�l��ݒ肷��B
            //
            private void setMouse(int x,int y){
                Dimension d = getSize(null);
                setMouseX(d,x);
                setMouseY(d,y);
                main_mouse_view.drawVal();
            }

            // �}�E�X��X���W���X�l���v�Z����B
            //
            private void setMouseX(Dimension d,int x){

                //�w�����W�v�Z
                float val;
                Float x_max = new Float(gr_x_length);
                float min = 0.0f;
                float max = x_max.floatValue();

                //�r�w�k����
                val = xPosConv(d.width,view_rec.width,min,max,x);
                x_length_mouse = val;

            }

            // �}�E�X��Y���W���Y�l���v�Z����B
            //
            private void setMouseY(Dimension d,int y){

                //�x�����W�v�Z
                float val;
                Float s_min;
                float min;
                Float s_max;
                float max;

                //���C���q�[�^�[�P���x
                s_min = new Float(main1_h_t_min_pro);
                min   = s_min.floatValue();
                s_max = new Float(main1_h_t_max_pro);
                max   = s_max.floatValue();
                val = yPosConv(d.height,view_rec.height,min,max,y);
                y_main1_h_t_mouse = val;

                //���C���q�[�^�[�P���x�v���t�@�C��
                s_min = new Float(main1_h_t_pf_min_pro);
                min   = s_min.floatValue();
                s_max = new Float(main1_h_t_pf_max_pro);
                max   = s_max.floatValue();
                val = yPosConv(d.height,view_rec.height,min,max,y);
                y_main1_h_t_pf_mouse = val;

                //���a
                s_min = new Float(dia_min_pro);
                min   = s_min.floatValue();
                s_max = new Float(dia_max_pro);
                max   = s_max.floatValue();
                val = yPosConv(d.height,view_rec.height,min,max,y);
                y_dia_mouse = val;

                //�����グ���x
                s_min = new Float(sxl_st_min_pro);
                min   = s_min.floatValue();
                s_max = new Float(sxl_st_max_pro);
                max   = s_max.floatValue();
                val = yPosConv(d.height,view_rec.height,min,max,y);
                y_sxl_st_mouse = val;

                //�����グ���x�v���t�@�C��
                s_min = new Float(sxl_st_pf_min_pro);
                min   = s_min.floatValue();
                s_max = new Float(sxl_st_pf_max_pro);
                max   = s_max.floatValue();
                val = yPosConv(d.height,view_rec.height,min,max,y);
                y_sxl_st_pf_mouse = val;
            }

            // �}�E�X��X�l���O���t��̂w���W�����߂�
            //
            private float xPos(int d_width,int v_width,float min,float max,float val){
                float x_dot = (float)v_width / (max - min);
                float x = x_dot * (val - min);
                return x;
            }

            // �w���W���l�����߂�
            //
            private float xPosConv(int d_width,int v_width,float min,float max,int x){
                float x_dot = (float)v_width / (max - min);
                float val = x / x_dot + min;
                return val;
            }

            // �}�E�X��Y�l���O���t��̂x���W�����߂�
            //
            private float yPos(int d_height,int v_height,float min,float max,float val){
                float y_dot = (float)v_height / (max - min);
                float y = (float)d_height - y_dot * (val - min);
                return y;
            }

            // �x���W���l�����߂�
            //
            private float yPosConv(int d_height,int v_height,float min,float max,int y){
                float y_dot = (float)v_height / (max - min);
                float val = (d_height - y) / y_dot + min;
                return val;
            }

            /***********************************************
             * �}�E�X����̏���
             *
             ***********************************************/
            class MainViewMouseMotion implements MouseMotionListener {

                //
                public void mouseDragged(MouseEvent e){

                }

                // �}�E�X�ړ���Listener
                //  �}�E�X�ʒu�iX,Y���W�j���f�[�^�l��ύX����B
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
     *       �w���̖ڐ��\���p�p�l��
     *
     *******************************************************/
    public class XSc extends JScrollPane {

        private Rectangle       view_rec        = null;
        private View            view            = null;

        // ---------- �R���X�g���N�^ -----------------------
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

        // X���̕\��
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
         * X���̖ڐ��\���p�l��
         ***************************************************/
        class View extends JPanel {
            Rectangle view_rec = null;

            View(){
                super();
                setName("View");
                setLayout(null);
                setBackground(BACK_COL);
            }

            // �\���g��ݒ肷��B
            //
            public void setViewRec(Rectangle rec){
                view_rec = rec;
            }

            // X���ڐ���`�悷��B
            //
            public void paint(Graphics g){
                Dimension d = getSize(null);

                g.setColor(BACK_COL);
                g.fillRect(0,0,d.width,d.height);
                drawMemLine(g);
                drawMem(g);
            }

            // �ڐ�����`�悷��B
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

            // �ڐ���`�悷��B
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
     *       �x���O���t�����ڐ�
     *
     *******************************************************/
    public class Y1Sc extends JScrollPane {

        private Rectangle   view_rec    = null;
        private View        view        = null;

        // ---------- �R���X�g���N�^ -----------------------
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

        // Y���������ĕ`�悷��B
        //
        public void chgYSize(){
            view.repaint();
        }

        /***************************************************
         * Y�������̖ڐ���\������
         ***************************************************/
        class View extends JPanel {
            Rectangle view_rec = null;

            // ---------- �R���X�g���N�^ -------------------
            View(){
                super();
                setName("View");
                setLayout(null);
                setBackground(BACK_COL);
            }

            // �̈��ݒ肷��B
            //
            public void setViewRec(Rectangle rec){
                view_rec = rec;
            }

            // �ڐ���`�悷��B
            //
            public void paint(Graphics g){
                Dimension d = getSize(null);
                g.setColor(BACK_COL);
                g.fillRect(0,0,d.width,d.height);

                g.setColor(MEM_LINE1_COL);
                drawMemLine(g);
                drawMemString(g);
            }

            // �ڐ���`�悷��
            //
            private void drawMemString(Graphics g){ 
                Dimension d = getSize(null);

                //���C���q�[�^�[�P���x
                g.setColor(MAIN1_H_T_COL);
                drawMem(g,d,45,main1_h_t_min_pro,main1_h_t_max_pro);

                //���C���q�[�^�[�P���x�v���t�@�C��
                g.setColor(MAIN1_H_T_PF_COL);
                drawMem(g,d,35,main1_h_t_pf_min_pro,main1_h_t_pf_max_pro);

                //���a
                g.setColor(DIA_COL);
                drawMem(g,d,25,dia_min_pro,dia_max_pro);

                //�����グ���x
                g.setColor(SXL_ST_COL);
                drawMem(g,d,15,sxl_st_min_pro,sxl_st_max_pro);

                //�����グ���x�v���t�@�C��
                g.setColor(SXL_ST_PF_COL);
                drawMem(g,d,5,sxl_st_pf_min_pro,sxl_st_pf_max_pro);
            }

            // �ڐ�����`�悷��
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

            // �ڐ���`�悷��
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
     *       �x���O���t�E���p�l��
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

        // �ĕ`�悷��B
        //
        public void chgYSize(){
            view.repaint();
        }

        /***************************************************
         * Y���E����\������
         ***************************************************/
        class View extends JPanel {
            Rectangle view_rec = null;

            // ---------- �R���X�g���N�^ -------------------
            View(){
                super();
                setName("View");
                setLayout(null);
                setBackground(BACK_COL);
            }

            // �̈��ݒ肷��B
            //
            public void setViewRec(Rectangle rec){
                view_rec = rec;
            }

            // �ڐ�����`�悷��B
            //
            public void paint(Graphics g){
                Dimension d = getSize(null);
                g.setColor(BACK_COL);
                g.fillRect(0,0,d.width,d.height);
                drawMemLine(g);
            }

            // �ڐ�����`�悷��B
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
     * �O���t�\���̈��Listener
     *
     *******************************************************/
    class PVGrEventCompo implements ComponentListener {

        private JPanel main_view    = null;     // �O���t�\���p�l��
        private JPanel x_view       = null;     // X���ڐ��p�l��
        private JPanel y1_view      = null;     // Y�����ڐ��p�l��
        private JPanel y2_view      = null;     // Y���E�p�l��

        // ---------- �R���X�g���N�^ -----------------------
        PVGrEventCompo(){

        }

        //
        // �O���t�\���p�l����ێ�����
        public void setMainView(JPanel view){
            main_view = view;
        }

        //
        // X���ڐ��\���p�l����ێ�����
        public void setXView(JPanel view){
            x_view = view;
        }

        //
        // Y�������ڐ��\���p�l����ێ�����
        public void setY1View(JPanel view){
            y1_view = view;
        }

        //
        // Y���E���\���p�l����ێ�����
        public void setY2View(JPanel view){
            y2_view = view;
        }

        //
        // �ړ����̏���
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
