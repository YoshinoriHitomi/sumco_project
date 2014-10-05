package cz;

import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.awt.event.WindowAdapter;
import java.awt.event.WindowEvent;
import java.util.Locale;

import javax.swing.ButtonGroup;
import javax.swing.JButton;
import javax.swing.JDialog;
import javax.swing.JLabel;
import javax.swing.JOptionPane;
import javax.swing.JRadioButton;
import javax.swing.JTextField;
import javax.swing.plaf.metal.MetalBorders.Flush3DBorder;
import javax.swing.text.AttributeSet;
import javax.swing.text.BadLocationException;
import javax.swing.text.Document;
import javax.swing.text.PlainDocument;

import czclass.CZNativeHikiage;
import czclass.CZParamHikiage;

/*******************************************************************************
 * �����グ��ݒ�Window
 * @author  (KPK Co.,Ltd.)
 * @version 1.0 (2003/04/01)
 * @@ T6�ǉ�
 *******************************************************************************/
public class CZBtSetWin extends JDialog {

    private final int Y = 30;

    private JLabel          lab[]       = new JLabel[16];
    private ROBTText        ro_bt       = null;
    private PGDIText        pgid        = null;
    private HINSYUText      hinsyu      = null;
    private HOUIText        houi        = null;
    private TYPEText        type        = null;
    private ROWText         row         = null;
    private OIText          oi          = null;
    private GAPText         gap         = null;
    private RUTUBOText      rutubo      = null;
    private PULLARText      pullar      = null;
    private TOPARText       topar       = null;
    private KEIText         kei         = null;
    private LENGText        pleng       = null;
    private SHIKOMIText     shikomi     = null;
    private SHIKOMIRText    shikomir    = null;
    private ZANEKIText      zaneki      = null;

    private JLabel          t_lab[]     = new JLabel[6];    //@@ 5 -> 6
    private TText           ttext[]     = new TText[6];     //@@ 5 -> 6

    private JLabel          proc_lab    = null;
    private JRadioButton    procPad[]   = new JRadioButton[3];

    private JButton         endt_button = null;
    private TText           end_ttext   = null;

    private JButton     start_button    = null;
    private JButton     restart_button  = null;
    private JButton     cancel_button   = null;

    //_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
    // �R���X�g���N�^
    //_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
    CZBtSetWin(){
        super();

        setTitle("�����グ�����ݒ�");
        setSize(360,820);
        setResizable(false);
        setModal(true);

        addWindowListener(
            new WindowAdapter(){
                public void windowClosing(WindowEvent e){
                    setDefault();
                }
        });

        getContentPane().setLayout(null);
        // ����n�Q�Ƌ@�\    @20131021
        if(CZSystemDefine.REFERENCE_RUN == CZSystem.getRunLevel()){
            getContentPane().setBackground(CZSystemDefine.DEFAULT_REFERENCE_BACKGROUND_COL);
        }else{
            getContentPane().setBackground(CZSystemDefine.DEFAULT_BACKGROUND_COL);
        }

        int i  = 0;
        int x1 = 20;
        int x2 = 130;
        int y  = 20;
        lab[i] = new JLabel("BtNo",JLabel.CENTER);
        lab[i].setBounds(x1, y, 100, 24);
        lab[i].setLocale(new Locale("ja","JP"));
        lab[i].setFont(new java.awt.Font("dialog", 0, 16));
        lab[i].setBorder(new Flush3DBorder());
        lab[i].setForeground(java.awt.Color.black);
        getContentPane().add(lab[i]);

        ro_bt = new ROBTText();
        ro_bt.setBounds(x2, y, 200, 24);
        getContentPane().add(ro_bt);

        i++;
        y+=Y;
        lab[i] = new JLabel("PGID",JLabel.CENTER);
        lab[i].setBounds(x1, y, 100, 24);
        lab[i].setLocale(new Locale("ja","JP"));
        lab[i].setFont(new java.awt.Font("dialog", 0, 16));
        lab[i].setBorder(new Flush3DBorder());
        lab[i].setForeground(java.awt.Color.black);
        getContentPane().add(lab[i]);

        pgid = new PGDIText();
        pgid.setBounds(x2, y, 200, 24);
        getContentPane().add(pgid);

        i++;
        y+=Y;
        lab[i] = new JLabel("�i��",JLabel.CENTER);
        lab[i].setBounds(x1, y, 100, 24);
        lab[i].setLocale(new Locale("ja","JP"));
        lab[i].setFont(new java.awt.Font("dialog", 0, 16));
        lab[i].setBorder(new Flush3DBorder());
        lab[i].setForeground(java.awt.Color.black);
        getContentPane().add(lab[i]);

        hinsyu = new HINSYUText();
        hinsyu.setBounds(x2, y, 200, 24);
        getContentPane().add(hinsyu);

        i++;
        y+=Y;
        lab[i] = new JLabel("����",JLabel.CENTER);
        lab[i].setBounds(x1, y, 100, 24);
        lab[i].setLocale(new Locale("ja","JP"));
        lab[i].setFont(new java.awt.Font("dialog", 0, 16));
        lab[i].setBorder(new Flush3DBorder());
        lab[i].setForeground(java.awt.Color.black);
        getContentPane().add(lab[i]);

        houi = new HOUIText();
        houi.setBounds(x2, y, 80, 24);
        getContentPane().add(houi);

        i++;
        y+=Y;
        lab[i] = new JLabel("�^�C�v",JLabel.CENTER);
        lab[i].setBounds(x1, y, 100, 24);
        lab[i].setLocale(new Locale("ja","JP"));
        lab[i].setFont(new java.awt.Font("dialog", 0, 16));
        lab[i].setBorder(new Flush3DBorder());
        lab[i].setForeground(java.awt.Color.black);
        getContentPane().add(lab[i]);

        type = new TYPEText();
        type.setBounds(x2, y, 80, 24);
        getContentPane().add(type);

        i++;
        y+=Y;
        lab[i] = new JLabel("���R",JLabel.CENTER);
        lab[i].setBounds(x1, y, 100, 24);
        lab[i].setLocale(new Locale("ja","JP"));
        lab[i].setFont(new java.awt.Font("dialog", 0, 16));
        lab[i].setBorder(new Flush3DBorder());
        lab[i].setForeground(java.awt.Color.black);
        getContentPane().add(lab[i]);

        row = new ROWText();
        row.setBounds(x2, y, 200, 24);
        getContentPane().add(row);

        i++;
        y+=Y;
        lab[i] = new JLabel("�_�f",JLabel.CENTER);
        lab[i].setBounds(x1, y, 100, 24);
        lab[i].setLocale(new Locale("ja","JP"));
        lab[i].setFont(new java.awt.Font("dialog", 0, 16));
        lab[i].setBorder(new Flush3DBorder());
        lab[i].setForeground(java.awt.Color.black);
        getContentPane().add(lab[i]);

        oi = new OIText();
        oi.setBounds(x2, y, 200, 24);
        getContentPane().add(oi);

        i++;
        y+=Y;
        lab[i] = new JLabel("GAP",JLabel.CENTER);
        lab[i].setBounds(x1, y, 100, 24);
        lab[i].setLocale(new Locale("ja","JP"));
        lab[i].setFont(new java.awt.Font("dialog", 0, 16));
        lab[i].setBorder(new Flush3DBorder());
        lab[i].setForeground(java.awt.Color.black);
        getContentPane().add(lab[i]);

        gap = new GAPText();
        gap.setBounds(x2, y, 80, 24);
        getContentPane().add(gap);

        i++;
        y+=Y;
        lab[i] = new JLabel("���c�{",JLabel.CENTER);
        lab[i].setBounds(x1, y, 100, 24);
        lab[i].setLocale(new Locale("ja","JP"));
        lab[i].setFont(new java.awt.Font("dialog", 0, 16));
        lab[i].setBorder(new Flush3DBorder());
        lab[i].setForeground(java.awt.Color.black);
        getContentPane().add(lab[i]);

        rutubo = new RUTUBOText();
        rutubo.setBounds(x2, y, 80, 24);
        getContentPane().add(rutubo);

        i++;
        y+=Y;
        lab[i] = new JLabel("�v�� Ar",JLabel.CENTER);
        lab[i].setBounds(x1, y, 100, 24);
        lab[i].setLocale(new Locale("ja","JP"));
        lab[i].setFont(new java.awt.Font("dialog", 0, 16));
        lab[i].setBorder(new Flush3DBorder());
        lab[i].setForeground(java.awt.Color.black);
        getContentPane().add(lab[i]);

        pullar = new PULLARText();
        pullar.setBounds(x2, y, 80, 24);
        getContentPane().add(pullar);

        i++;
        y+=Y;
        lab[i] = new JLabel("�g�b�v Ar",JLabel.CENTER);
        lab[i].setBounds(x1, y, 100, 24);
        lab[i].setLocale(new Locale("ja","JP"));
        lab[i].setFont(new java.awt.Font("dialog", 0, 16));
        lab[i].setBorder(new Flush3DBorder());
        lab[i].setForeground(java.awt.Color.black);
        getContentPane().add(lab[i]);

        topar = new TOPARText();
        topar.setBounds(x2, y, 80, 24);
        getContentPane().add(topar);

        i++;
        y+=Y;
        lab[i] = new JLabel("���a",JLabel.CENTER);
        lab[i].setBounds(x1, y, 100, 24);
        lab[i].setLocale(new Locale("ja","JP"));
        lab[i].setFont(new java.awt.Font("dialog", 0, 16));
        lab[i].setBorder(new Flush3DBorder());
        lab[i].setForeground(java.awt.Color.black);
        getContentPane().add(lab[i]);

        kei = new KEIText();
        kei.setBounds(x2, y, 80, 24);
        getContentPane().add(kei);

        i++;
        y+=Y;
        lab[i] = new JLabel("���㒷",JLabel.CENTER);
        lab[i].setBounds(x1, y, 100, 24);
        lab[i].setLocale(new Locale("ja","JP"));
        lab[i].setFont(new java.awt.Font("dialog", 0, 16));
        lab[i].setBorder(new Flush3DBorder());
        lab[i].setForeground(java.awt.Color.black);
        getContentPane().add(lab[i]);

        pleng = new LENGText();
        pleng.setBounds(x2, y, 80, 24);
        getContentPane().add(pleng);

        i++;
        y+=Y;
        lab[i] = new JLabel("�d���� (I)",JLabel.CENTER);
        lab[i].setBounds(x1, y, 100, 24);
        lab[i].setLocale(new Locale("ja","JP"));
        lab[i].setFont(new java.awt.Font("dialog", 0, 16));
        lab[i].setBorder(new Flush3DBorder());
        lab[i].setForeground(java.awt.Color.black);
        getContentPane().add(lab[i]);

        shikomi = new SHIKOMIText();
        shikomi.setBounds(x2, y, 80, 24);
        getContentPane().add(shikomi);

        i++;
        y+=Y;
        lab[i] = new JLabel("�d���� (R)",JLabel.CENTER);
        lab[i].setBounds(x1, y, 100, 24);
        lab[i].setLocale(new Locale("ja","JP"));
        lab[i].setFont(new java.awt.Font("dialog", 0, 16));
        lab[i].setBorder(new Flush3DBorder());
        lab[i].setForeground(java.awt.Color.black);
        getContentPane().add(lab[i]);

        shikomir = new SHIKOMIRText();
        shikomir.setBounds(x2, y, 80, 24);
        getContentPane().add(shikomir);

        i++;
        y+=Y;
        lab[i] = new JLabel("�c�t",JLabel.CENTER);
        lab[i].setBounds(x1, y, 100, 24);
        lab[i].setLocale(new Locale("ja","JP"));
        lab[i].setFont(new java.awt.Font("dialog", 0, 16));
        lab[i].setBorder(new Flush3DBorder());
        lab[i].setForeground(java.awt.Color.black);
        getContentPane().add(lab[i]);

        zaneki = new ZANEKIText();
        zaneki.setBounds(x2, y, 80, 24);
        getContentPane().add(zaneki);

        i  = 0;
        y+=Y;
        t_lab[i] = new JLabel("�n�� (T1)",JLabel.CENTER);
        t_lab[i].setBounds(x1, y, 100, 24);
        t_lab[i].setLocale(new Locale("ja","JP"));
        t_lab[i].setFont(new java.awt.Font("dialog", 0, 16));
        t_lab[i].setBorder(new Flush3DBorder());
        t_lab[i].setForeground(java.awt.Color.black);
        getContentPane().add(t_lab[i]);

        ttext[i] = new TText();
        ttext[i].setBounds(x2, y, 80, 24);
        getContentPane().add(ttext[i]);

        i++;
        y+=Y;
        t_lab[i] = new JLabel("���� (T2)",JLabel.CENTER);
        t_lab[i].setBounds(x1, y, 100, 24);
        t_lab[i].setLocale(new Locale("ja","JP"));
        t_lab[i].setFont(new java.awt.Font("dialog", 0, 16));
        t_lab[i].setBorder(new Flush3DBorder());
        t_lab[i].setForeground(java.awt.Color.black);
        getContentPane().add(t_lab[i]);

        ttext[i] = new TText();
        ttext[i].setBounds(x2, y, 80, 24);
        getContentPane().add(ttext[i]);

        i++;
        y+=Y;
        t_lab[i] = new JLabel("��] (T3)",JLabel.CENTER);
        t_lab[i].setBounds(x1, y, 100, 24);
        t_lab[i].setLocale(new Locale("ja","JP"));
        t_lab[i].setFont(new java.awt.Font("dialog", 0, 16));
        t_lab[i].setBorder(new Flush3DBorder());
        t_lab[i].setForeground(java.awt.Color.black);
        getContentPane().add(t_lab[i]);

        ttext[i] = new TText();
        ttext[i].setBounds(x2, y, 80, 24);
        getContentPane().add(ttext[i]);

        i++;
        y+=Y;
        t_lab[i] = new JLabel("��o (T4)",JLabel.CENTER);
        t_lab[i].setBounds(x1, y, 100, 24);
        t_lab[i].setLocale(new Locale("ja","JP"));
        t_lab[i].setFont(new java.awt.Font("dialog", 0, 16));
        t_lab[i].setBorder(new Flush3DBorder());
        t_lab[i].setForeground(java.awt.Color.black);
        getContentPane().add(t_lab[i]);

        ttext[i] = new TText();
        ttext[i].setBounds(x2, y, 80, 24);
        getContentPane().add(ttext[i]);

        i++;
        y+=Y;
        t_lab[i] = new JLabel("���� (T5)",JLabel.CENTER);
        t_lab[i].setBounds(x1, y, 100, 24);
        t_lab[i].setLocale(new Locale("ja","JP"));
        t_lab[i].setFont(new java.awt.Font("dialog", 0, 16));
        t_lab[i].setBorder(new Flush3DBorder());
        t_lab[i].setForeground(java.awt.Color.black);
        getContentPane().add(t_lab[i]);

        ttext[i] = new TText();
        ttext[i].setBounds(x2, y, 80, 24);
        getContentPane().add(ttext[i]);

        i++;
        y+=Y;
        t_lab[i] = new JLabel("�萔 (T6)",JLabel.CENTER);   //@@
        t_lab[i].setBounds(x1, y, 100, 24);
        t_lab[i].setLocale(new Locale("ja","JP"));
        t_lab[i].setFont(new java.awt.Font("dialog", 0, 16));
        t_lab[i].setBorder(new Flush3DBorder());
        t_lab[i].setForeground(java.awt.Color.black);
        getContentPane().add(t_lab[i]);

        ttext[i] = new TText();
        ttext[i].setBounds(x2, y, 80, 24);
        getContentPane().add(ttext[i]);

        i++;
        y+=Y;
        endt_button = new JButton("��o (T4) �đ�");
        endt_button.setBounds(x1, y, 150, 24);
        endt_button.setLocale(new Locale("ja","JP"));
        endt_button.setFont(new java.awt.Font("dialog", 0, 16));
        endt_button.setBorder(new Flush3DBorder());
        endt_button.setForeground(java.awt.Color.black);
        endt_button.addActionListener(new EndReSend());
        getContentPane().add(endt_button);

        end_ttext = new TText();
        end_ttext.setBounds(x2+50, y, 80, 24);
        getContentPane().add(end_ttext);

        y+=Y;
        proc_lab = new JLabel("�X�^�[�g�v���Z�X",JLabel.CENTER);
        proc_lab.setBounds(x1, y, 150, 24);
        proc_lab.setLocale(new Locale("ja","JP"));
        proc_lab.setFont(new java.awt.Font("dialog", 0, 16));
        proc_lab.setBorder(new Flush3DBorder());
        proc_lab.setForeground(java.awt.Color.black);
        getContentPane().add(proc_lab);

        ButtonGroup group = new ButtonGroup();
        i=0;
        x2 = 50;
        procPad[i] = new JRadioButton(CZSystem.getProcName(CZSystemDefine.VAC));
        procPad[i].setMnemonic('V');
        procPad[i].setBorder(new Flush3DBorder());
        procPad[i].setBounds(x1+(x2*i)+160, y, x2, 24);
        group.add(procPad[i]);
        getContentPane().add(procPad[i]);

        i++;
        procPad[i] = new JRadioButton(CZSystem.getProcName(CZSystemDefine.MELT));
        procPad[i].setMnemonic('M');
        procPad[i].setSelected(true);       
        procPad[i].setBorder(new Flush3DBorder());
        procPad[i].setBounds(x1+(x2*i)+160, y, x2, 24);
        group.add(procPad[i]);
        getContentPane().add(procPad[i]);

        i++;
        procPad[i] = new JRadioButton(CZSystem.getProcName(CZSystemDefine.DIP));
        procPad[i].setMnemonic('D');
        procPad[i].setBorder(new Flush3DBorder());
        procPad[i].setBounds(x1+(x2*i)+160, y, x2, 24);
        group.add(procPad[i]);
        getContentPane().add(procPad[i]);

        y+=Y;
        x2 = 80;
        start_button = new JButton(" �J  �n ");
        start_button.setBounds(x1, y, x2, 24);
        start_button.setLocale(new Locale("ja","JP"));
        start_button.setFont(new java.awt.Font("dialog", 0, 16));
        start_button.setBorder(new Flush3DBorder());
        start_button.setForeground(java.awt.Color.black);
        start_button.addActionListener(new StartSend());
        getContentPane().add(start_button);

        restart_button = new JButton(" ��  �J ");
        restart_button.setBounds(x1+x2+20, y, x2, 24);
        restart_button.setLocale(new Locale("ja","JP"));
        restart_button.setFont(new java.awt.Font("dialog", 0, 16));
        restart_button.setBorder(new Flush3DBorder());
        restart_button.setForeground(java.awt.Color.black);
        restart_button.addActionListener(new ReStartSend());
        getContentPane().add(restart_button);

        cancel_button = new JButton(" �I  �� ");
        cancel_button.setBounds(x1+x2+x2+70, y, x2, 24);
        cancel_button.setLocale(new Locale("ja","JP"));
        cancel_button.setFont(new java.awt.Font("dialog", 0, 16));
        cancel_button.setBorder(new Flush3DBorder());
        cancel_button.setForeground(java.awt.Color.black);
        cancel_button.addActionListener(new Cancel());
        getContentPane().add(cancel_button);
    }

    //
    //
    //
    public boolean setDefault(){
//@@        CZSystem.log("CZBtSetWin","setDefault() Start (Set Current Hikiage)");

        CZNativeHikiage tbl = CZSystem.getBtSet();
        if (null == tbl) return false;
        // �o�b�`No
        ro_bt.setText(tbl.getBatch());

        // PG-ID
        pgid.setText(tbl.getPgid());

        // �i��
        hinsyu.setText(tbl.getHinshu());

        // ����
        houi.setText(tbl.getHoui());

        // �^�C�v
        type.setText(tbl.getH_type());

        // ���R
        row.setText(tbl.getHiteikou());

        // �_�f
        oi.setText(tbl.getSanso());

        // GAP
        gap.setText(tbl.getGap());

        // ���c�{�a
        rutubo.setText(String.valueOf(tbl.getRutubo_kei()));

        // �v���A���S��
        pullar.setText(String.valueOf(tbl.getPull_ar()));

        // �g�b�v�A���S��
        topar.setText(String.valueOf(tbl.getTop_ar()));

        // ���a
        kei.setText(String.valueOf(tbl.getChokkei()));

        // ���㒷
        pleng.setText(String.valueOf(tbl.getHikiage_cho()));

        // �����d����
        shikomi.setText(String.valueOf(tbl.getI_sikomi()));

        // �ǉ��d����
        shikomir.setText(String.valueOf(tbl.getT_sikomi()));

        // �c�t��
        zaneki.setText(String.valueOf(tbl.getZaneki()));

        // ���V�s�[�m���i�n���j
        ttext[0].setText(String.valueOf(tbl.getNo_youkai()  ));

        // ���V�s�[�m���i����j
        ttext[1].setText(String.valueOf(tbl.getNo_hikiage() ));

        // ���V�s�[�m���i��]�j
        ttext[2].setText(String.valueOf(tbl.getNo_kaiten()  ));

        // ���V�s�[�m���i��o�j
        ttext[3].setText(String.valueOf(tbl.getNo_toridasi()));

        // ���V�s�[�m���i���́j
        ttext[4].setText(String.valueOf(tbl.getNo_aturyoku()));

        // ���V�s�[�m���i�萔
        ttext[5].setText(String.valueOf(tbl.getNo_teisu()));

        return true;
    }

    //
    //
    //
    public boolean chgTimes(){
        int     val = 0;
        return true;
    }

    //
    //
    //
    private boolean errorMsg(Object msg[]){
        JOptionPane.showMessageDialog(null,msg,"�����グ�����ݒ���̓G���[",JOptionPane.ERROR_MESSAGE);
        return true;
    }

    //
    //
    //
    private int getStartProc(){

        if(procPad[0].isSelected()) return CZSystemDefine.VAC;
        if(procPad[1].isSelected()) return CZSystemDefine.MELT;
        if(procPad[2].isSelected()) return CZSystemDefine.DIP;

        return -1;
    }


    //
    //
    //
    public CZParamHikiage setParamT4(){
//@@            CZSystem.log("CZBtSetWin ","setParamT4()");

            CZParamHikiage ret = setParam();

            int no = -1;
            String t_no = end_ttext.getText();

            try{
                no = Integer.valueOf(t_no).intValue();  

                if(1 > no){
                    Object msg[] = {"��o (T4) �đ�",
                                    "���͂��������Ă��������I�I",
                                    "�e�[�u��No 1 �����ł�"};
                    errorMsg(msg);
                    return null;
                }   

                if(999 < no){
                    Object msg[] = {"��o (T4) �đ�",
                                    "���͂��������Ă��������I�I",
                                    "�e�[�u��No 1000 �ȏ�ł�"};
                    errorMsg(msg);
                    return null;
                }   
                ret.setNo_toridasi(no);
            }
            catch(Exception e){
                Object msg[] = {"��o (T4) �đ�",
                                "���͂��������Ă��������I�I"};
                errorMsg(msg);
                return null;
            }
            return ret;
        }

        //
        //
        //
        public CZParamHikiage setParam(){
//@@            CZSystem.log("CZBtSetWin","setParam()");

            CZParamHikiage ret = new CZParamHikiage();
            String ro_tmp = null;
            String bt_tmp = null;

            int no = -1;
            String tmp = null;

            /* �FBtNo */
            try{
                tmp = ro_bt.getText();

                if(null == tmp){
                    Object msg[] = {"Bt No",
                                    "���͂��������Ă��������I�I",
                                    "NULL"};
                    errorMsg(msg);
                    return null;
                }   

                if(1 > tmp.length()){
                    Object msg[] = {"Bt No",
                                    "���͂��������Ă��������I�I",
                                    "LENGTH"};
                    errorMsg(msg);
                    return null;
                }   

                ret.setBatch(new String(tmp));
            }
            catch(Exception e){
//@@                CZSystem.log("CZBtSetWin",""+ e);
                Object msg[] = {"Bt No",
                                "���͂��������Ă��������I�I",
                                "EXCEPTION"};
                errorMsg(msg);
                return null;
            }

            /* PGID */
            try{
                tmp = pgid.getText();

                if(null == tmp){
                    Object msg[] = {"PGID",
                                    "���͂��������Ă��������I�I"};
                    errorMsg(msg);
                    return null;
                }   

                if(1 > tmp.length()){
                    Object msg[] = {"PGID",
                                    "���͂��������Ă��������I�I"};
                    errorMsg(msg);
                    return null;
                }   

                ret.setPgid(new String(tmp));
            }
            catch(Exception e){
                Object msg[] = {"PGID",
                                "���͂��������Ă��������I�I"};
                errorMsg(msg);
                return null;
            }
            

            /* �i�� */
            try{
                tmp = hinsyu.getText();

                if(null == tmp){
                    Object msg[] = {"�i��",
                                    "���͂��������Ă��������I�I"};
                    errorMsg(msg);
                    return null;
                }   

                if(1 > tmp.length()){
                    Object msg[] = {"�i��",
                                    "���͂��������Ă��������I�I"};
                    errorMsg(msg);
                    return null;
                }   
                ret.setHinshu(new String(tmp));
            }
            catch(Exception e){
                Object msg[] = {"�i��",
                                "���͂��������Ă��������I�I"};
                errorMsg(msg);
                return null;
            }

            /* ���� */
            try{
                tmp = houi.getText();

                if(null == tmp){
                    Object msg[] = {"����",
                                    "���͂��������Ă��������I�I"};
                    errorMsg(msg);
                    return null;
                }   

                if(1 > tmp.length()){
                    Object msg[] = {"����",
                                    "���͂��������Ă��������I�I"};
                    errorMsg(msg);
                    return null;
                }   

                ret.setHoui(new String(tmp));
            }
            catch(Exception e){
                Object msg[] = {"����",
                                "���͂��������Ă��������I�I"};
                errorMsg(msg);
                return null;
            }

            /* �^�C�v */
            try{
                tmp = type.getText();

                if(null == tmp){
                    Object msg[] = {"�^�C�v",
                                    "���͂��������Ă��������I�I"};
                    errorMsg(msg);
                    return null;
                }   

                if(1 > tmp.length()){
                    Object msg[] = {"�^�C�v",
                                    "���͂��������Ă��������I�I"};
                    errorMsg(msg);
                    return null;
                }   

                //Send
                ret.setH_type(new String(tmp));
            }
            catch(Exception e){
                Object msg[] = {"�^�C�v",
                                "���͂��������Ă��������I�I"};
                errorMsg(msg);
                return null;
            }


            /* ���R */
            try{
                tmp = row.getText();

                if(null == tmp){
                    Object msg[] = {"���R",
                                    "���͂��������Ă��������I�I"};
                    errorMsg(msg);
                    return null;
                }   

                if(1 > tmp.length()){
                    Object msg[] = {"���R",
                                    "���͂��������Ă��������I�I"};
                    errorMsg(msg);
                    return null;
                }   

                //Send
                ret.setHiteikou(new String(tmp));
            }
            catch(Exception e){
                Object msg[] = {"���R",
                                "���͂��������Ă��������I�I"};
                errorMsg(msg);
                return null;
            }


            /* �_�f */
            try{
                tmp = oi.getText();

                if(null == tmp){
                    Object msg[] = {"�_�f",
                                    "���͂��������Ă��������I�I"};
                    errorMsg(msg);
                    return null;
                }   

                if(1 > tmp.length()){
                    Object msg[] = {"�_�f",
                                    "���͂��������Ă��������I�I"};
                    errorMsg(msg);
                    return null;
                }   

                //Send
                ret.setSanso(new String(tmp));
            }
            catch(Exception e){
                Object msg[] = {"�_�f",
                                "���͂��������Ă��������I�I"};
                errorMsg(msg);
                return null;
            }


            /* GAP */
            try{
                tmp = gap.getText();

                if(null == tmp){
                    Object msg[] = {"GAP",
                                    "���͂��������Ă��������I�I"};
                    errorMsg(msg);
                    return null;
                }   

                if(1 > tmp.length()){
                    Object msg[] = {"GAP",
                                    "���͂��������Ă��������I�I"};
                    errorMsg(msg);
                    return null;
                }   

                //Send
                ret.setGap(new String(tmp));
            }
            catch(Exception e){
                Object msg[] = {"GAP",
                                "���͂��������Ă��������I�I"};
                errorMsg(msg);
                return null;
            }


            /* ���c�{ */
            try{
                tmp = rutubo.getText();
                no = Integer.valueOf(tmp).intValue();   

                //Send
                ret.setRutubo_kei(no);
            }
            catch(Exception e){
                Object msg[] = {"���c�{",
                                "���͂��������Ă��������I�I"};
                errorMsg(msg);
                return null;
            }


            /* �v�� Ar */
            try{
                tmp = pullar.getText();
                no = Integer.valueOf(tmp).intValue();   

                //Send
                ret.setPull_ar(no);
            }
            catch(Exception e){
                Object msg[] = {"�v�� Ar",
                                "���͂��������Ă��������I�I"};
                errorMsg(msg);
                return null;
            }

            /* �g�b�v Ar */
            try{
                tmp = topar.getText();
                no = Integer.valueOf(tmp).intValue();   

                //Send
                ret.setTop_ar(no);
            }
            catch(Exception e){
                Object msg[] = {"�g�b�v Ar",
                                "���͂��������Ă��������I�I"};
                errorMsg(msg);
                return null;
            }

            /* ���a */
            try{
                tmp = kei.getText();
                no = Integer.valueOf(tmp).intValue();   

                //Send
                ret.setChokkei(no);
            }
            catch(Exception e){
                Object msg[] = {"���a",
                                "���͂��������Ă��������I�I"};
                errorMsg(msg);
                return null;
            }

            /* ���㒷 */
            try{
                tmp = pleng.getText();
                no = Integer.valueOf(tmp).intValue();   

                //Send
                ret.setHikiage_cho(no);
            }
            catch(Exception e){
                Object msg[] = {"���㒷",
                                "���͂��������Ă��������I�I"};
                errorMsg(msg);
                return null;
            }

            /* �d����(I) */
            try{
                tmp = shikomi.getText();
                no = Integer.valueOf(tmp).intValue();   

                //Send
                ret.setI_sikomi(no);
            }
            catch(Exception e){
                Object msg[] = {"�d����(I)",
                                "���͂��������Ă��������I�I"};
                errorMsg(msg);
                return null;
            }

            /* �d����(R) */
            try{
                tmp = shikomir.getText();
                no = Integer.valueOf(tmp).intValue();   

                //Send
                ret.setT_sikomi(no);
            }
            catch(Exception e){
                Object msg[] = {"�d����(R)",
                                "���͂��������Ă��������I�I"};
                errorMsg(msg);
                return null;
            }

            /* �c�t */
            try{
                tmp = zaneki.getText();
                no = Integer.valueOf(tmp).intValue();   

                //Send
                ret.setZaneki(no);
            }
            catch(Exception e){
                Object msg[] = {"�c�t",
                                "���͂��������Ă��������I�I"};
                errorMsg(msg);
                return null;
            }

            /************/
            /* �n��(T1) */
            try{
                tmp = ttext[0].getText();
                no = Integer.valueOf(tmp).intValue();   

                if(1 > no){
                    Object msg[] = {"�n��(T1)",
                                    "���͂��������Ă��������I�I",
                                    "�e�[�u��No 1 �����ł�"};
                    errorMsg(msg);
                    return null;
                }   

                if(999 < no){
                    Object msg[] = {"�n��(T1)",
                                    "���͂��������Ă��������I�I",
                                    "�e�[�u��No 1000 �ȏ�ł�"};
                    errorMsg(msg);
                    return null;
                }   

                ret.setNo_youkai(no);
            }
            catch(Exception e){
                Object msg[] = {"�n��(T1)",
                                "���͂��������Ă��������I�I"};
                errorMsg(msg);
                return null;
            }

            /* ����(T2) */
            try{
                tmp = ttext[1].getText();
                no = Integer.valueOf(tmp).intValue();   

                if(1 > no){
                    Object msg[] = {"����(T2)",
                                    "���͂��������Ă��������I�I",
                                    "�e�[�u��No 1 �����ł�"};
                    errorMsg(msg);
                    return null;
                }   

                if(999 < no){
                    Object msg[] = {"����(T2)",
                                    "���͂��������Ă��������I�I",
                                    "�e�[�u��No 1000 �ȏ�ł�"};
                    errorMsg(msg);
                    return null;
                }   

                ret.setNo_hikiage(no);
            }
            catch(Exception e){
                Object msg[] = {"����(T2)",
                                "���͂��������Ă��������I�I"};
                errorMsg(msg);
                return null;
            }

            /* ��](T3) */
            try{
                tmp = ttext[2].getText();
                no = Integer.valueOf(tmp).intValue();   

                if(1 > no){
                    Object msg[] = {"��](T3)",
                                    "���͂��������Ă��������I�I",
                                    "�e�[�u��No 1 �����ł�"};
                    errorMsg(msg);
                    return null;
                }   

                if(999 < no){
                    Object msg[] = {"��](T3)",
                                    "���͂��������Ă��������I�I",
                                    "�e�[�u��No 1000 �ȏ�ł�"};
                    errorMsg(msg);
                    return null;
                }   

                ret.setNo_kaiten(no);
            }
            catch(Exception e){
                Object msg[] = {"��](T3)",
                        "���͂��������Ă��������I�I"};
                errorMsg(msg);
                return null;
            }

            /* ��o(T4) */
            try{
                tmp = ttext[3].getText();
                no = Integer.valueOf(tmp).intValue();   

                if(1 > no){
                    Object msg[] = {"��o(T4)",
                                    "���͂��������Ă��������I�I",
                                    "�e�[�u��No 1 �����ł�"};
                    errorMsg(msg);
                    return null;
                }   

                if(999 < no){
                    Object msg[] = {"��o(T4)",
                                    "���͂��������Ă��������I�I",
                                    "�e�[�u��No 1000 �ȏ�ł�"};
                    errorMsg(msg);
                    return null;
                }   

                ret.setNo_toridasi(no);
            }
            catch(Exception e){
                Object msg[] = {"��o(T4)",
                                "���͂��������Ă��������I�I"};
                errorMsg(msg);
                return null;
            }

            /* ����(T5) */
            try{
                tmp = ttext[4].getText();
                no = Integer.valueOf(tmp).intValue();   

                if(1 > no){
                    Object msg[] = {"����(T5)",
                                    "���͂��������Ă��������I�I",
                                    "�e�[�u��No 1 �����ł�"};
                    errorMsg(msg);
                    return null;
                }   

                if(999 < no){
                    Object msg[] = {"����(T5)",
                                    "���͂��������Ă��������I�I",
                                    "�e�[�u��No 1000 �ȏ�ł�"};
                    errorMsg(msg);
                    return null;
                }   

                ret.setNo_aturyoku(no);
            }
            catch(Exception e){
                Object msg[] = {"����(T5)",
                                "���͂��������Ă��������I�I"};
                errorMsg(msg);
                return null;
            }

            /* @@ �萔(T6) */
            try{
                tmp = ttext[5].getText();
                no = Integer.valueOf(tmp).intValue();   

                if(1 > no){
                    Object msg[] = {"�萔(T6)",
                                    "���͂��������Ă��������I�I",
                                    "�e�[�u��No 1 �����ł�"};
                    errorMsg(msg);
                    return null;
                }   

                if(999 < no){
                    Object msg[] = {"�萔(T6)",
                                    "���͂��������Ă��������I�I",
                                    "�e�[�u��No 1000 �ȏ�ł�"};
                    errorMsg(msg);
                    return null;
                }   

                ret.setNo_teisu(no);
            }
            catch(Exception e){
                Object msg[] = {"�萔(T6)",
                                "���͂��������Ă��������I�I"};
                errorMsg(msg);
                return null;
            }
            return ret;
       }


        /***********************************************************************
         *
         *   ���o���e�[�u���̍đ�
         *
         ***********************************************************************/
        class EndReSend implements ActionListener {
            public void actionPerformed(ActionEvent ev){
                CZSystem.log("CZBtSetWin EndReSend","��o (T4) �đ�");

                CZParamHikiage ret = null;

                ret = setParamT4();

                if(null == ret){
                    CZSystem.log("CZBtSetWin EndReSend","��o (T4) �đ��G���[");
                    return;
                }

                if(!CZSystem.CZOperateToridasi(ret)){
                    Object msg[] = {"�o�^�o���܂���ł����i���o���e�[�u���̍đ��j",
                                    "���͂��������Ă��������I�I"};
                    errorMsg(msg);
                }
                return;
            }
        }

        /***********************************************************************
         *
         *   �����グ�������M(�J�n)
         *
         ***********************************************************************/
        class StartSend implements ActionListener {
            public void actionPerformed(ActionEvent ev){
                CZSystem.log("CZBtSetWin StartSend","�J�n");

                CZParamHikiage ret = null;

                ret = setParam();
                if(null == ret){
                    CZSystem.log("CZBtSetWin StartSend","�J�n�G���[");
                    return;
                }

                ret.setPno_start(getStartProc());

                ret.setP_kaisi(CZSystemDefine.START_PROC_START);

                if(!CZSystem.CZOperateHikiage(ret)){
                    Object msg[] = {"�o�^�o���܂���ł���",
                                    "���͂��������Ă��������I�I"};
                    errorMsg(msg);
                }
                return;
            }
        }

        /***********************************************************************
         *
         *   �����グ�������M(�ĊJ)
         *
         ***********************************************************************/
        class ReStartSend implements ActionListener {
            public void actionPerformed(ActionEvent ev){
                CZSystem.log("CZBtSetWin ReStartSend","�ĊJ");

                CZParamHikiage ret = null;

                ret = setParam();
                if(null == ret){
                    CZSystem.log("CZBtSetWin ReStartSend","�ĊJ�G���[");
                    return;
                }

                ret.setPno_start(getStartProc());
                ret.setP_kaisi(CZSystemDefine.START_PROC_RESTART);

                CZSystem.CZOperateHikiage(ret);
                return;
            }
        }


        /***********************************************************************
         *
         *   ��ʏ���
         *
         ***********************************************************************/
        class Cancel implements ActionListener {
            public void actionPerformed(ActionEvent ev){
                CZSystem.log("CZBtSetWin","Cancel");

                setVisible(false);
                setDefault();
                return;
            }
        }


        /***********************************************************************
         *
         *   PV�O���tMin,Max��ݒ肷��TextField
         *       ���l�݂̂��󂯕t����
         *
         ***********************************************************************/
        public class PVText extends JTextField {

            PVText(){
                super();
                setFont(new java.awt.Font("dialog", 0, 16));
            }

            protected Document createDefaultModel() {
                return new NumericDocument();
            }

            class NumericDocument extends PlainDocument {
                String validValues = "0123456789.-";

                public void insertString( int offset, String str, AttributeSet a )
                                                throws BadLocationException {

                    char[] val = str.toCharArray();
                    for (int i = 0; i < val.length; i++) {
                        if(validValues.indexOf(val[i]) == -1) return;
                    }

                    super.insertString( offset, str, a );
                    return ;
                }
            }
        }


        /***********************************************************************
         *
         *   PV�O���t�{����ݒ肷��TextField
         *       ���l�݂̂��󂯕t����
         *
         ***********************************************************************/
        public class TimesText extends JTextField {

            TimesText(){
                super();
                setFont(new java.awt.Font("dialog", 0, 16));
            }

            protected Document createDefaultModel() {
                return new NumericDocument();
            }

            class NumericDocument extends PlainDocument {
                String validValues = "0123456789";

                public void insertString( int offset, String str, AttributeSet a )
                                                    throws BadLocationException {

                    char[] val = str.toCharArray();
                    for (int i = 0; i < val.length; i++) {
                        if(validValues.indexOf(val[i]) == -1) return;
                    }
                    super.insertString( offset, str, a );

                    chgTimes();
                }

                public void remove(int offs, int len)
                                                throws BadLocationException {
                    super.remove(offs,len);
                    chgTimes();
                }
            }
        }


        /***********************************************************************
         *
         *   �F�ԁA�o�b�`No��ݒ肷��TextField
         *
         ***********************************************************************/
        public class ROBTText extends JTextField {

            ROBTText(){
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
                String validValues = "0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ-";

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

                //
                //
                public void remove(int offs, int len)
                                                throws BadLocationException {
                    super.remove(offs,len);
                    chgTimes();
                }
            }
        }



        /***********************************************************************
         *
         *   PGID��ݒ肷��TextField
         *
         ***********************************************************************/
        public class PGDIText extends JTextField {

            PGDIText(){
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
                String validValues = "0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ-.";

                //
                //
                public void insertString( int offset, String str, AttributeSet a )
                                                        throws BadLocationException {

                    if(7 < getLength()) return;
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
                    chgTimes();
                }
            }
        }


        /***********************************************************************
         *
         *   �i���ݒ肷��TextField
         *
         ***********************************************************************/
        public class HINSYUText extends JTextField {

            HINSYUText(){
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
                String validValues = "0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ-";

                //
                //
                public void insertString( int offset, String str, AttributeSet a )
                                                    throws BadLocationException {

                    if(11 < getLength()) return;
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
                    chgTimes();
                }
            }
        }

        /***********************************************************************
         *
         *   ���ʂ�ݒ肷��TextField
         *
         ***********************************************************************/
        public class HOUIText extends JTextField {

            HOUIText(){
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
                String validValues = "01245";

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
                chgTimes();
                }
            }
        }


        /***********************************************************************
         *
         *   �^�C�v��ݒ肷��TextField
         *
         ***********************************************************************/
        public class TYPEText extends JTextField {

            TYPEText(){
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
                String validValues = "PNBSb";

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
                    chgTimes();
            }
        }
    }


    /***********************************************************************
     *
     *   ���R��ݒ肷��TextField
     *
     ***********************************************************************/
    public class ROWText extends JTextField {

        ROWText(){
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
            String validValues = "0123456789.-";

            //
            //
            public void insertString( int offset, String str, AttributeSet a )
                                                throws BadLocationException {

                if(14 < getLength()) return;
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
                chgTimes();
            }
        }
    }

    /***********************************************************************
     *
     *   �_�f��ݒ肷��TextField
     *
     ***********************************************************************/
    public class OIText extends JTextField {

        OIText(){
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

            //
            //
            public void remove(int offs, int len)
                                                throws BadLocationException {
                super.remove(offs,len);
                chgTimes();
            }
        }
    }

    /***********************************************************************
     *
     *   GAP��ݒ肷��TextField
     *
     ***********************************************************************/
    public class GAPText extends JTextField {

        GAPText(){
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
                chgTimes();
            }
        }
    }


    /***********************************************************************
     *
     *   ���c�{�a�i�C���`�j��ݒ肷��TextField
     *
     ***********************************************************************/
    public class RUTUBOText extends JTextField {

        RUTUBOText(){
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
                chgTimes();
            }
        }
    }

    /***********************************************************************
     *
     *   �v���A���S����ݒ肷��TextField
     *
     ***********************************************************************/
    public class PULLARText extends JTextField {

        PULLARText(){
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
                chgTimes();
            }
        }
    }

    /***********************************************************************
     *
     *   �g�b�v�A���S����ݒ肷��TextField
     *
     ***********************************************************************/
    public class TOPARText extends JTextField {

        TOPARText(){
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
                chgTimes();
            }
        }
    }

    /***********************************************************************
     *
     *   �����グ�a(mm)��ݒ肷��TextField
     *
     ***********************************************************************/
    public class KEIText extends JTextField {

        KEIText(){
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
                chgTimes();
            }
        }
    }

    /***********************************************************************
     *
     *   �����グ������ݒ肷��TextField
     *
     ***********************************************************************/
    public class LENGText extends JTextField {

        LENGText(){
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


                if(3 < getLength()) return;
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
                chgTimes();
            }
        }
    }


    /***********************************************************************
     *
     *   �d���݃C�j�V������ݒ肷��TextField
     *
     ***********************************************************************/
    public class SHIKOMIText extends JTextField {

        SHIKOMIText(){
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

                if(5 < getLength()) return;
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
                chgTimes();
            }
        }
    }


    /***********************************************************************
     *
     *   �d���ݒǉ���ݒ肷��TextField
     *
     ***********************************************************************/
    public class SHIKOMIRText extends JTextField {

        SHIKOMIRText(){
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

                if(5 < getLength()) return;
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
                chgTimes();
            }
        }
    }


    /***********************************************************************
    *
    *   �c�t��ݒ肷��TextField
    *
    ***********************************************************************/
    public class ZANEKIText extends JTextField {

        ZANEKIText(){
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

                if(5 < getLength()) return;
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
                chgTimes();
            }
        }
    }


    /***************************************************************************
     *
     *   ����e�[�u������͂���TextField
     *
     ***************************************************************************/
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
                chgTimes();
            }
        }
    }
}
