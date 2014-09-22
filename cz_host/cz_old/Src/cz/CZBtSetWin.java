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
 * 引き上げ条設定Window
 * @author  (KPK Co.,Ltd.)
 * @version 1.0 (2003/04/01)
 * @@ T6追加
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
    // コンストラクタ
    //_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
    CZBtSetWin(){
        super();

        setTitle("引き上げ条件設定");
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
        // 他基地参照機能    @20131021
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
        lab[i] = new JLabel("品種",JLabel.CENTER);
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
        lab[i] = new JLabel("方位",JLabel.CENTER);
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
        lab[i] = new JLabel("タイプ",JLabel.CENTER);
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
        lab[i] = new JLabel("比抵抗",JLabel.CENTER);
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
        lab[i] = new JLabel("酸素",JLabel.CENTER);
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
        lab[i] = new JLabel("ルツボ",JLabel.CENTER);
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
        lab[i] = new JLabel("プル Ar",JLabel.CENTER);
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
        lab[i] = new JLabel("トップ Ar",JLabel.CENTER);
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
        lab[i] = new JLabel("直径",JLabel.CENTER);
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
        lab[i] = new JLabel("引上長",JLabel.CENTER);
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
        lab[i] = new JLabel("仕込み (I)",JLabel.CENTER);
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
        lab[i] = new JLabel("仕込み (R)",JLabel.CENTER);
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
        lab[i] = new JLabel("残液",JLabel.CENTER);
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
        t_lab[i] = new JLabel("溶解 (T1)",JLabel.CENTER);
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
        t_lab[i] = new JLabel("引上 (T2)",JLabel.CENTER);
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
        t_lab[i] = new JLabel("回転 (T3)",JLabel.CENTER);
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
        t_lab[i] = new JLabel("取出 (T4)",JLabel.CENTER);
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
        t_lab[i] = new JLabel("圧力 (T5)",JLabel.CENTER);
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
        t_lab[i] = new JLabel("定数 (T6)",JLabel.CENTER);   //@@
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
        endt_button = new JButton("取出 (T4) 再送");
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
        proc_lab = new JLabel("スタートプロセス",JLabel.CENTER);
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
        start_button = new JButton(" 開  始 ");
        start_button.setBounds(x1, y, x2, 24);
        start_button.setLocale(new Locale("ja","JP"));
        start_button.setFont(new java.awt.Font("dialog", 0, 16));
        start_button.setBorder(new Flush3DBorder());
        start_button.setForeground(java.awt.Color.black);
        start_button.addActionListener(new StartSend());
        getContentPane().add(start_button);

        restart_button = new JButton(" 再  開 ");
        restart_button.setBounds(x1+x2+20, y, x2, 24);
        restart_button.setLocale(new Locale("ja","JP"));
        restart_button.setFont(new java.awt.Font("dialog", 0, 16));
        restart_button.setBorder(new Flush3DBorder());
        restart_button.setForeground(java.awt.Color.black);
        restart_button.addActionListener(new ReStartSend());
        getContentPane().add(restart_button);

        cancel_button = new JButton(" 終  了 ");
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
        // バッチNo
        ro_bt.setText(tbl.getBatch());

        // PG-ID
        pgid.setText(tbl.getPgid());

        // 品種
        hinsyu.setText(tbl.getHinshu());

        // 方位
        houi.setText(tbl.getHoui());

        // タイプ
        type.setText(tbl.getH_type());

        // 比抵抗
        row.setText(tbl.getHiteikou());

        // 酸素
        oi.setText(tbl.getSanso());

        // GAP
        gap.setText(tbl.getGap());

        // ルツボ径
        rutubo.setText(String.valueOf(tbl.getRutubo_kei()));

        // プルアルゴン
        pullar.setText(String.valueOf(tbl.getPull_ar()));

        // トップアルゴン
        topar.setText(String.valueOf(tbl.getTop_ar()));

        // 直径
        kei.setText(String.valueOf(tbl.getChokkei()));

        // 引上長
        pleng.setText(String.valueOf(tbl.getHikiage_cho()));

        // 初期仕込量
        shikomi.setText(String.valueOf(tbl.getI_sikomi()));

        // 追加仕込量
        shikomir.setText(String.valueOf(tbl.getT_sikomi()));

        // 残液量
        zaneki.setText(String.valueOf(tbl.getZaneki()));

        // レシピーＮｏ（溶解）
        ttext[0].setText(String.valueOf(tbl.getNo_youkai()  ));

        // レシピーＮｏ（引上）
        ttext[1].setText(String.valueOf(tbl.getNo_hikiage() ));

        // レシピーＮｏ（回転）
        ttext[2].setText(String.valueOf(tbl.getNo_kaiten()  ));

        // レシピーＮｏ（取出）
        ttext[3].setText(String.valueOf(tbl.getNo_toridasi()));

        // レシピーＮｏ（圧力）
        ttext[4].setText(String.valueOf(tbl.getNo_aturyoku()));

        // レシピーＮｏ（定数
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
        JOptionPane.showMessageDialog(null,msg,"引き上げ条件設定入力エラー",JOptionPane.ERROR_MESSAGE);
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
                    Object msg[] = {"取出 (T4) 再送",
                                    "入力を見直してください！！",
                                    "テーブルNo 1 未満です"};
                    errorMsg(msg);
                    return null;
                }   

                if(999 < no){
                    Object msg[] = {"取出 (T4) 再送",
                                    "入力を見直してください！！",
                                    "テーブルNo 1000 以上です"};
                    errorMsg(msg);
                    return null;
                }   
                ret.setNo_toridasi(no);
            }
            catch(Exception e){
                Object msg[] = {"取出 (T4) 再送",
                                "入力を見直してください！！"};
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

            /* 炉BtNo */
            try{
                tmp = ro_bt.getText();

                if(null == tmp){
                    Object msg[] = {"Bt No",
                                    "入力を見直してください！！",
                                    "NULL"};
                    errorMsg(msg);
                    return null;
                }   

                if(1 > tmp.length()){
                    Object msg[] = {"Bt No",
                                    "入力を見直してください！！",
                                    "LENGTH"};
                    errorMsg(msg);
                    return null;
                }   

                ret.setBatch(new String(tmp));
            }
            catch(Exception e){
//@@                CZSystem.log("CZBtSetWin",""+ e);
                Object msg[] = {"Bt No",
                                "入力を見直してください！！",
                                "EXCEPTION"};
                errorMsg(msg);
                return null;
            }

            /* PGID */
            try{
                tmp = pgid.getText();

                if(null == tmp){
                    Object msg[] = {"PGID",
                                    "入力を見直してください！！"};
                    errorMsg(msg);
                    return null;
                }   

                if(1 > tmp.length()){
                    Object msg[] = {"PGID",
                                    "入力を見直してください！！"};
                    errorMsg(msg);
                    return null;
                }   

                ret.setPgid(new String(tmp));
            }
            catch(Exception e){
                Object msg[] = {"PGID",
                                "入力を見直してください！！"};
                errorMsg(msg);
                return null;
            }
            

            /* 品種 */
            try{
                tmp = hinsyu.getText();

                if(null == tmp){
                    Object msg[] = {"品種",
                                    "入力を見直してください！！"};
                    errorMsg(msg);
                    return null;
                }   

                if(1 > tmp.length()){
                    Object msg[] = {"品種",
                                    "入力を見直してください！！"};
                    errorMsg(msg);
                    return null;
                }   
                ret.setHinshu(new String(tmp));
            }
            catch(Exception e){
                Object msg[] = {"品種",
                                "入力を見直してください！！"};
                errorMsg(msg);
                return null;
            }

            /* 方位 */
            try{
                tmp = houi.getText();

                if(null == tmp){
                    Object msg[] = {"方位",
                                    "入力を見直してください！！"};
                    errorMsg(msg);
                    return null;
                }   

                if(1 > tmp.length()){
                    Object msg[] = {"方位",
                                    "入力を見直してください！！"};
                    errorMsg(msg);
                    return null;
                }   

                ret.setHoui(new String(tmp));
            }
            catch(Exception e){
                Object msg[] = {"方位",
                                "入力を見直してください！！"};
                errorMsg(msg);
                return null;
            }

            /* タイプ */
            try{
                tmp = type.getText();

                if(null == tmp){
                    Object msg[] = {"タイプ",
                                    "入力を見直してください！！"};
                    errorMsg(msg);
                    return null;
                }   

                if(1 > tmp.length()){
                    Object msg[] = {"タイプ",
                                    "入力を見直してください！！"};
                    errorMsg(msg);
                    return null;
                }   

                //Send
                ret.setH_type(new String(tmp));
            }
            catch(Exception e){
                Object msg[] = {"タイプ",
                                "入力を見直してください！！"};
                errorMsg(msg);
                return null;
            }


            /* 比抵抗 */
            try{
                tmp = row.getText();

                if(null == tmp){
                    Object msg[] = {"比抵抗",
                                    "入力を見直してください！！"};
                    errorMsg(msg);
                    return null;
                }   

                if(1 > tmp.length()){
                    Object msg[] = {"比抵抗",
                                    "入力を見直してください！！"};
                    errorMsg(msg);
                    return null;
                }   

                //Send
                ret.setHiteikou(new String(tmp));
            }
            catch(Exception e){
                Object msg[] = {"比抵抗",
                                "入力を見直してください！！"};
                errorMsg(msg);
                return null;
            }


            /* 酸素 */
            try{
                tmp = oi.getText();

                if(null == tmp){
                    Object msg[] = {"酸素",
                                    "入力を見直してください！！"};
                    errorMsg(msg);
                    return null;
                }   

                if(1 > tmp.length()){
                    Object msg[] = {"酸素",
                                    "入力を見直してください！！"};
                    errorMsg(msg);
                    return null;
                }   

                //Send
                ret.setSanso(new String(tmp));
            }
            catch(Exception e){
                Object msg[] = {"酸素",
                                "入力を見直してください！！"};
                errorMsg(msg);
                return null;
            }


            /* GAP */
            try{
                tmp = gap.getText();

                if(null == tmp){
                    Object msg[] = {"GAP",
                                    "入力を見直してください！！"};
                    errorMsg(msg);
                    return null;
                }   

                if(1 > tmp.length()){
                    Object msg[] = {"GAP",
                                    "入力を見直してください！！"};
                    errorMsg(msg);
                    return null;
                }   

                //Send
                ret.setGap(new String(tmp));
            }
            catch(Exception e){
                Object msg[] = {"GAP",
                                "入力を見直してください！！"};
                errorMsg(msg);
                return null;
            }


            /* ルツボ */
            try{
                tmp = rutubo.getText();
                no = Integer.valueOf(tmp).intValue();   

                //Send
                ret.setRutubo_kei(no);
            }
            catch(Exception e){
                Object msg[] = {"ルツボ",
                                "入力を見直してください！！"};
                errorMsg(msg);
                return null;
            }


            /* プル Ar */
            try{
                tmp = pullar.getText();
                no = Integer.valueOf(tmp).intValue();   

                //Send
                ret.setPull_ar(no);
            }
            catch(Exception e){
                Object msg[] = {"プル Ar",
                                "入力を見直してください！！"};
                errorMsg(msg);
                return null;
            }

            /* トップ Ar */
            try{
                tmp = topar.getText();
                no = Integer.valueOf(tmp).intValue();   

                //Send
                ret.setTop_ar(no);
            }
            catch(Exception e){
                Object msg[] = {"トップ Ar",
                                "入力を見直してください！！"};
                errorMsg(msg);
                return null;
            }

            /* 直径 */
            try{
                tmp = kei.getText();
                no = Integer.valueOf(tmp).intValue();   

                //Send
                ret.setChokkei(no);
            }
            catch(Exception e){
                Object msg[] = {"直径",
                                "入力を見直してください！！"};
                errorMsg(msg);
                return null;
            }

            /* 引上長 */
            try{
                tmp = pleng.getText();
                no = Integer.valueOf(tmp).intValue();   

                //Send
                ret.setHikiage_cho(no);
            }
            catch(Exception e){
                Object msg[] = {"引上長",
                                "入力を見直してください！！"};
                errorMsg(msg);
                return null;
            }

            /* 仕込み(I) */
            try{
                tmp = shikomi.getText();
                no = Integer.valueOf(tmp).intValue();   

                //Send
                ret.setI_sikomi(no);
            }
            catch(Exception e){
                Object msg[] = {"仕込み(I)",
                                "入力を見直してください！！"};
                errorMsg(msg);
                return null;
            }

            /* 仕込み(R) */
            try{
                tmp = shikomir.getText();
                no = Integer.valueOf(tmp).intValue();   

                //Send
                ret.setT_sikomi(no);
            }
            catch(Exception e){
                Object msg[] = {"仕込み(R)",
                                "入力を見直してください！！"};
                errorMsg(msg);
                return null;
            }

            /* 残液 */
            try{
                tmp = zaneki.getText();
                no = Integer.valueOf(tmp).intValue();   

                //Send
                ret.setZaneki(no);
            }
            catch(Exception e){
                Object msg[] = {"残液",
                                "入力を見直してください！！"};
                errorMsg(msg);
                return null;
            }

            /************/
            /* 溶解(T1) */
            try{
                tmp = ttext[0].getText();
                no = Integer.valueOf(tmp).intValue();   

                if(1 > no){
                    Object msg[] = {"溶解(T1)",
                                    "入力を見直してください！！",
                                    "テーブルNo 1 未満です"};
                    errorMsg(msg);
                    return null;
                }   

                if(999 < no){
                    Object msg[] = {"溶解(T1)",
                                    "入力を見直してください！！",
                                    "テーブルNo 1000 以上です"};
                    errorMsg(msg);
                    return null;
                }   

                ret.setNo_youkai(no);
            }
            catch(Exception e){
                Object msg[] = {"溶解(T1)",
                                "入力を見直してください！！"};
                errorMsg(msg);
                return null;
            }

            /* 引上(T2) */
            try{
                tmp = ttext[1].getText();
                no = Integer.valueOf(tmp).intValue();   

                if(1 > no){
                    Object msg[] = {"引上(T2)",
                                    "入力を見直してください！！",
                                    "テーブルNo 1 未満です"};
                    errorMsg(msg);
                    return null;
                }   

                if(999 < no){
                    Object msg[] = {"引上(T2)",
                                    "入力を見直してください！！",
                                    "テーブルNo 1000 以上です"};
                    errorMsg(msg);
                    return null;
                }   

                ret.setNo_hikiage(no);
            }
            catch(Exception e){
                Object msg[] = {"引上(T2)",
                                "入力を見直してください！！"};
                errorMsg(msg);
                return null;
            }

            /* 回転(T3) */
            try{
                tmp = ttext[2].getText();
                no = Integer.valueOf(tmp).intValue();   

                if(1 > no){
                    Object msg[] = {"回転(T3)",
                                    "入力を見直してください！！",
                                    "テーブルNo 1 未満です"};
                    errorMsg(msg);
                    return null;
                }   

                if(999 < no){
                    Object msg[] = {"回転(T3)",
                                    "入力を見直してください！！",
                                    "テーブルNo 1000 以上です"};
                    errorMsg(msg);
                    return null;
                }   

                ret.setNo_kaiten(no);
            }
            catch(Exception e){
                Object msg[] = {"回転(T3)",
                        "入力を見直してください！！"};
                errorMsg(msg);
                return null;
            }

            /* 取出(T4) */
            try{
                tmp = ttext[3].getText();
                no = Integer.valueOf(tmp).intValue();   

                if(1 > no){
                    Object msg[] = {"取出(T4)",
                                    "入力を見直してください！！",
                                    "テーブルNo 1 未満です"};
                    errorMsg(msg);
                    return null;
                }   

                if(999 < no){
                    Object msg[] = {"取出(T4)",
                                    "入力を見直してください！！",
                                    "テーブルNo 1000 以上です"};
                    errorMsg(msg);
                    return null;
                }   

                ret.setNo_toridasi(no);
            }
            catch(Exception e){
                Object msg[] = {"取出(T4)",
                                "入力を見直してください！！"};
                errorMsg(msg);
                return null;
            }

            /* 圧力(T5) */
            try{
                tmp = ttext[4].getText();
                no = Integer.valueOf(tmp).intValue();   

                if(1 > no){
                    Object msg[] = {"圧力(T5)",
                                    "入力を見直してください！！",
                                    "テーブルNo 1 未満です"};
                    errorMsg(msg);
                    return null;
                }   

                if(999 < no){
                    Object msg[] = {"圧力(T5)",
                                    "入力を見直してください！！",
                                    "テーブルNo 1000 以上です"};
                    errorMsg(msg);
                    return null;
                }   

                ret.setNo_aturyoku(no);
            }
            catch(Exception e){
                Object msg[] = {"圧力(T5)",
                                "入力を見直してください！！"};
                errorMsg(msg);
                return null;
            }

            /* @@ 定数(T6) */
            try{
                tmp = ttext[5].getText();
                no = Integer.valueOf(tmp).intValue();   

                if(1 > no){
                    Object msg[] = {"定数(T6)",
                                    "入力を見直してください！！",
                                    "テーブルNo 1 未満です"};
                    errorMsg(msg);
                    return null;
                }   

                if(999 < no){
                    Object msg[] = {"定数(T6)",
                                    "入力を見直してください！！",
                                    "テーブルNo 1000 以上です"};
                    errorMsg(msg);
                    return null;
                }   

                ret.setNo_teisu(no);
            }
            catch(Exception e){
                Object msg[] = {"定数(T6)",
                                "入力を見直してください！！"};
                errorMsg(msg);
                return null;
            }
            return ret;
       }


        /***********************************************************************
         *
         *   取り出しテーブルの再送
         *
         ***********************************************************************/
        class EndReSend implements ActionListener {
            public void actionPerformed(ActionEvent ev){
                CZSystem.log("CZBtSetWin EndReSend","取出 (T4) 再送");

                CZParamHikiage ret = null;

                ret = setParamT4();

                if(null == ret){
                    CZSystem.log("CZBtSetWin EndReSend","取出 (T4) 再送エラー");
                    return;
                }

                if(!CZSystem.CZOperateToridasi(ret)){
                    Object msg[] = {"登録出来ませんでした（取り出しテーブルの再送）",
                                    "入力を見直してください！！"};
                    errorMsg(msg);
                }
                return;
            }
        }

        /***********************************************************************
         *
         *   引き上げ条件送信(開始)
         *
         ***********************************************************************/
        class StartSend implements ActionListener {
            public void actionPerformed(ActionEvent ev){
                CZSystem.log("CZBtSetWin StartSend","開始");

                CZParamHikiage ret = null;

                ret = setParam();
                if(null == ret){
                    CZSystem.log("CZBtSetWin StartSend","開始エラー");
                    return;
                }

                ret.setPno_start(getStartProc());

                ret.setP_kaisi(CZSystemDefine.START_PROC_START);

                if(!CZSystem.CZOperateHikiage(ret)){
                    Object msg[] = {"登録出来ませんでした",
                                    "入力を見直してください！！"};
                    errorMsg(msg);
                }
                return;
            }
        }

        /***********************************************************************
         *
         *   引き上げ条件送信(再開)
         *
         ***********************************************************************/
        class ReStartSend implements ActionListener {
            public void actionPerformed(ActionEvent ev){
                CZSystem.log("CZBtSetWin ReStartSend","再開");

                CZParamHikiage ret = null;

                ret = setParam();
                if(null == ret){
                    CZSystem.log("CZBtSetWin ReStartSend","再開エラー");
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
         *   画面消去
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
         *   PVグラフMin,Maxを設定するTextField
         *       数値のみを受け付ける
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
         *   PVグラフ倍率を設定するTextField
         *       数値のみを受け付ける
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
         *   炉番、バッチNoを設定するTextField
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
         *   PGIDを設定するTextField
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
         *   品種を設定するTextField
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
         *   方位を設定するTextField
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
         *   タイプを設定するTextField
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
     *   比抵抗を設定するTextField
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
     *   酸素を設定するTextField
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
     *   GAPを設定するTextField
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
     *   ルツボ径（インチ）を設定するTextField
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
     *   プルアルゴンを設定するTextField
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
     *   トップアルゴンを設定するTextField
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
     *   引き上げ径(mm)を設定するTextField
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
     *   引き上げ長さを設定するTextField
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
     *   仕込みイニシャルを設定するTextField
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
     *   仕込み追加を設定するTextField
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
    *   残液を設定するTextField
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
     *   制御テーブルを入力するTextField
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
