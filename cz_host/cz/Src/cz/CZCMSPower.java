package cz;

import java.awt.Color;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.util.Locale;

import javax.swing.ButtonGroup;
import javax.swing.ImageIcon;
import javax.swing.JButton;
import javax.swing.JDialog;
import javax.swing.JLabel;
import javax.swing.JRadioButton;
import javax.swing.plaf.metal.MetalBorders.Flush3DBorder;

import czclass.CZNativeDengen;

/**
 * PV表示項目設定Window
 * @author  (KPK Co.,Ltd.)
 * @version 1.0 (2003/04/01)
 *
 */

public class CZCMSPower extends JDialog {
    public final int POWER_COUNT = 10 ;

    public final int POWER_NONE = 0 ;
    public final int POWER_OFF  = 1 ;
    public final int POWER_ON   = 2 ;

    public final Color COLOR_NONE = CZSystemDefine.DEFAULT_BACKGROUND_COL;
    public final Color COLOR_ON   = java.awt.Color.blue;
    public final Color COLOR_OFF  = java.awt.Color.red;

    public int send_status[] = new int[POWER_COUNT] ;

    private JButton     send_button   = null;
    private JButton     cancel_button = null;

    public static final String POWER_NAME[] = {
                                new String("シード昇降"),
                                new String("シード回転"),
                                new String("ルツボ昇降"),
                                new String("ルツボ回転"),
                                new String("結晶保持"),
                                new String("メインヒータ１"),
                                new String("メインヒータ２"),
                                new String("ボトムヒータ"),
                                new String("シードヒータ"),
                                new String("磁場")};

    private JLabel          unit_lab[]      = new JLabel[POWER_COUNT];
    private JButton         now_button[]    = new JButton[POWER_COUNT];

    private JRadioButton    pw_on[]     = new JRadioButton[POWER_COUNT];
    private JRadioButton    pw_off[]    = new JRadioButton[POWER_COUNT];
    private JRadioButton    pw_none[]   = new JRadioButton[POWER_COUNT];

    private UpdateThread    updateTh    = null;
    
    //
    //
    //
    CZCMSPower(){
        super();

        setTitle("電源");
        setSize(1060,230);
        setResizable(false);
        setModal(false);

        getContentPane().setLayout(null);
        // 他基地参照機能    @20131021
        if(CZSystemDefine.REFERENCE_RUN == CZSystem.getRunLevel()){
            getContentPane().setBackground(CZSystemDefine.DEFAULT_REFERENCE_BACKGROUND_COL);
        }else{
            getContentPane().setBackground(CZSystemDefine.DEFAULT_BACKGROUND_COL);
        }

        send_button = new JButton("実  行");
        send_button.setBounds(120, 160, 120, 24);
        send_button.setLocale(new Locale("ja","JP"));
        send_button.setFont(new java.awt.Font("dialog", 0, 18));
        send_button.setBorder(new Flush3DBorder());
        send_button.setForeground(java.awt.Color.black);
        send_button.addActionListener(new SendButton());
        getContentPane().add(send_button);

        cancel_button = new JButton("終  了");
        cancel_button.setBounds(900, 160, 120, 24);
        cancel_button.setLocale(new Locale("ja","JP"));
        cancel_button.setFont(new java.awt.Font("dialog", 0, 18));
        cancel_button.setBorder(new Flush3DBorder());
        cancel_button.setForeground(java.awt.Color.black);
        cancel_button.addActionListener(new CancelButton());
        getContentPane().add(cancel_button);

        JLabel lab1 = new JLabel("項    目",JLabel.CENTER);
        lab1.setBounds(20, 20, 100, 24);
        lab1.setLocale(new Locale("ja","JP"));
        lab1.setFont(new java.awt.Font("dialog", 0, 12));
        lab1.setBorder(new Flush3DBorder());
        lab1.setForeground(java.awt.Color.black);
        getContentPane().add(lab1);

        JLabel lab2 = new JLabel("電源状態",JLabel.CENTER);
        lab2.setBounds(20, 44, 100, 24);
        lab2.setLocale(new Locale("ja","JP"));
        lab2.setFont(new java.awt.Font("dialog", 0, 12));
        lab2.setBorder(new Flush3DBorder());
        lab2.setForeground(java.awt.Color.black);
        getContentPane().add(lab2);

        JLabel lab3 = new JLabel("電源状態変更",JLabel.CENTER);
        lab3.setBounds(20, 78, 100, 48);
        lab3.setLocale(new Locale("ja","JP"));
        lab3.setFont(new java.awt.Font("dialog", 0, 12));
        lab3.setBorder(new Flush3DBorder());
        lab3.setForeground(java.awt.Color.black);
        getContentPane().add(lab3);

        for(int i = 0 ; i < POWER_COUNT ; i++){
            createButton(i);
        }

        updateTh = new UpdateThread();
        updateTh.setPriority(Thread.MIN_PRIORITY);
        updateTh.start();
    }


    //
    //
    //
    private boolean setSendStatus(){

        try{
            for(int i = 0 ; i < POWER_COUNT ; i++){
                if(pw_none[i].isSelected()){
                    send_status[i] = POWER_NONE;
                    continue;
                }

                if(pw_off[i].isSelected()){
                    send_status[i] = POWER_OFF;
                    continue;
                }

                if(pw_on[i].isSelected()){
                    send_status[i] = POWER_ON;
                    continue;
                }
            } // for end
        }
        catch(Exception e){
            CZSystem.log("CZCMSPower",""+e);
            CZSystem.exit(-1,"CZCMSPower setSendStatus()");
            return false;
        }
        return true;
    }


    //
    //
    //
    private boolean createButton(int no){

        int width = 90;
        int x = no * width + 120;

        unit_lab[no] = new JLabel(POWER_NAME[no],JLabel.CENTER);
        unit_lab[no].setBounds(x, 20, width, 24);
        unit_lab[no].setLocale(new Locale("ja","JP"));
        unit_lab[no].setFont(new java.awt.Font("dialog", 0, 12));
        unit_lab[no].setBorder(new Flush3DBorder());
        unit_lab[no].setForeground(java.awt.Color.black);
        getContentPane().add(unit_lab[no]);

        now_button[no] = new JButton("ＯＦＦ");
        now_button[no].setBounds(x, 44, width, 24);
        now_button[no].setLocale(new Locale("ja","JP"));
        now_button[no].setFont(new java.awt.Font("dialog", 0, 12));
        now_button[no].setBorder(new Flush3DBorder());
        now_button[no].setForeground(java.awt.Color.white);
        now_button[no].setBackground(COLOR_NONE);
        now_button[no].addActionListener(new ResetButton(no));
        getContentPane().add(now_button[no]);

        ButtonGroup pw_group = new ButtonGroup();

        pw_on[no] = new JRadioButton("ON",new ImageIcon("images/rb.gif"));
        pw_on[no].setBounds(x+10, 78, width, 24);
        pw_on[no].setPressedIcon(new ImageIcon("images/rbp.gif"));
        pw_on[no].setRolloverIcon(new ImageIcon("images/rbr.gif"));
        pw_on[no].setRolloverSelectedIcon(new ImageIcon("images/rbrs.gif"));
        pw_on[no].setSelectedIcon(new ImageIcon("images/rbs.gif"));
        pw_on[no].setFocusPainted(false);
        pw_on[no].setBorderPainted(false);
        pw_on[no].setContentAreaFilled(false);
        pw_on[no].setSelected(false);
        getContentPane().add(pw_on[no]);
        pw_group.add(pw_on[no]);

        pw_off[no] = new JRadioButton("OFF",new ImageIcon("images/rb.gif"));
        pw_off[no].setBounds(x+10, 102, width, 24);
        pw_off[no].setPressedIcon(new ImageIcon("images/rbp.gif"));
        pw_off[no].setRolloverIcon(new ImageIcon("images/rbr.gif"));
        pw_off[no].setRolloverSelectedIcon(new ImageIcon("images/rbrs.gif"));
        pw_off[no].setSelectedIcon(new ImageIcon("images/rbs.gif"));
        pw_off[no].setFocusPainted(false);
        pw_off[no].setBorderPainted(false);
        pw_off[no].setContentAreaFilled(false);
        pw_off[no].setSelected(false);
        getContentPane().add(pw_off[no]);
        pw_group.add(pw_off[no]);

        pw_none[no] = new JRadioButton("NONE");
        pw_none[no].setBounds(x+10, 126, width, 24);
        pw_none[no].setVisible(false);
        pw_none[no].setSelected(false);
        getContentPane().add(pw_none[no]);
        pw_group.add(pw_none[no]);

        return true;
    }


    //
    //
    //
    public boolean setDefault(){

        CZSystem.log("CZCMSPower","setDefault()");
        for(int i = 0 ; i < POWER_COUNT ; i++)  pw_none[i].setSelected(true);

        //電源状態のセット
        setData();

        return true;
    }


    //
    //
    //
    private void setData(){

        CZNativeDengen dengen = CZSystem.getPowerStat();

        if(null == dengen){
            for(int i = 0 ; i < POWER_COUNT ; i++){
                now_button[i].setText("不明");
                now_button[i].setBackground(COLOR_ON);
            }
            return ;
        }

        int sw[] = dengen.getValue();

        //電源状態のセット
        for(int i = 0 ; i < POWER_COUNT ; i++){
            switch(sw[i]){
                case POWER_NONE :
                    now_button[i].setText("ＮＯＮＥ");
                    now_button[i].setBackground(COLOR_NONE);
                    break;
                case POWER_OFF  :
                    now_button[i].setText("ＯＦＦ");
                    now_button[i].setBackground(COLOR_OFF);
                    break;
                case POWER_ON   :
                    now_button[i].setText("ＯＮ");
                    now_button[i].setBackground(COLOR_ON);
                    break;

                default     :
                    now_button[i].setText("不明");
                    now_button[i].setBackground(COLOR_NONE);
                    break;
            }
        } // for end
    }

    /*
    *
    *
    *
    */
    class SendButton implements ActionListener {
        public void actionPerformed(ActionEvent ev){

            boolean ret = setSendStatus();

            CZSystem.log("CZCMSPower","SendButton actionPerformed");
            for(int i = 0 ; i <  POWER_COUNT ; i++){
                CZSystem.log("CZCMSPower","send_status [" + send_status[i] + "]");
            }
            System.out.println(" ");
            // Send
            if(ret){
                CZSystem.CZOperatePowerControl(send_status);
                send_button.setBackground(CZSystemDefine.BUTTON_SEND_COL);
            }
        }
    }


    /*
    *
    *
    *
    */
    class CancelButton implements ActionListener {
        public void actionPerformed(ActionEvent ev){
            setDefault();
            setVisible(false);
        }
    }

    /*
    *
    *
    *
    */
    class ResetButton implements ActionListener {
        private int my_no = -1;
            //
            //
            //
            ResetButton(int _no){
                super();
                my_no = _no;
            }

            public void actionPerformed(ActionEvent ev){
                CZSystem.log("CZCMSPower","ResetButton No[" + my_no + "]");
                pw_none[my_no].setSelected(true);
            }
        }


    /*
    *
    *
    *
    */
    class UpdateThread extends Thread {

        //
        //
        //
        UpdateThread(){
        }

        //
        //
        //
        public void run(){

            CZSystem.log("CZCMSPower","UpdateThread START");

            CZSystemQueue   que = new CZSystemQueue(10);
            CZEventAdapter  adp = new CZEventAdapter(que);
            CZEventSender.addCZEventListener(adp);

            while(true){
                try{
                    CZEventCL event = (CZEventCL)que.waitObject();

                    switch(event.getEvent()){
                        case CZEventCL.PV_RECEIVE :
                        case CZEventCL.RO_CHANGE  :
                        case CZEventCL.EV_F007    :     
                        case CZEventCL.EV_F009    :
                            setData();
                            break;
                    
                        case CZEventCL.EV_1031    :     
                            setDefault();
                            send_button.setBackground(CZSystemDefine.BUTTON_WAIT_COL);
                            break;
                        case CZEventCL.EV_8031    :
                            setDefault();
                            send_button.setBackground(CZSystemDefine.BUTTON_NORMAL_COL);
                            break;

                        default           : break;
                    } // switch end
                }
                catch(Exception e){

                }
            } // while end
        }
    }
}
