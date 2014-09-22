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

/*
 *   制御モード変更Window 
 * @author  (KPK Co.,Ltd.)
 * @version 1.0 (2003/04/01)
 *
 */
public class CZCMSControlMode extends JDialog {

    public final int MODE_NONE      = -1 ;
    public final int MODE_MANUAL    = CZSystemDefine.PROC_MANUAL;
    public final int MODE_AUTO      = CZSystemDefine.PROC_AUTO;

    public final Color COLOR_MANUAL = java.awt.Color.red;
    public final Color COLOR_AUTO   = java.awt.Color.blue;

    private int send_status         = MODE_NONE;
    private int now_mode            = -1;

    private JButton send_button     = null;
    private JButton cancel_button   = null;

    private JLabel  unit_lab        = null;
    private JButton now_button      = null;

    private JRadioButton mode_manual    = null;
    private JRadioButton mode_auto      = null;
    private JRadioButton mode_none      = null;

    private UpdateThread updateTh       = null;


    /*******************************************************
     *
     *******************************************************/
    CZCMSControlMode(){
        super();

        setTitle("制御モード");
        setSize(240,230);
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
        send_button.setBounds(20, 160, 100, 24);
        send_button.setLocale(new Locale("ja","JP"));
        send_button.setFont(new java.awt.Font("dialog", 0, 18));
        send_button.setBorder(new Flush3DBorder());
        send_button.setForeground(java.awt.Color.black);
        send_button.addActionListener(new SendButton());
        getContentPane().add(send_button);

        cancel_button = new JButton("終  了");
        cancel_button.setBounds(140, 160, 70, 24);
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

        JLabel lab2 = new JLabel("制御モード",JLabel.CENTER);
        lab2.setBounds(20, 44, 100, 24);
        lab2.setLocale(new Locale("ja","JP"));
        lab2.setFont(new java.awt.Font("dialog", 0, 12));
        lab2.setBorder(new Flush3DBorder());
        lab2.setForeground(java.awt.Color.black);
        getContentPane().add(lab2);

        JLabel lab3 = new JLabel("制御モード変更",JLabel.CENTER);
        lab3.setBounds(20, 78, 100, 48);
        lab3.setLocale(new Locale("ja","JP"));
        lab3.setFont(new java.awt.Font("dialog", 0, 12));
        lab3.setBorder(new Flush3DBorder());
        lab3.setForeground(java.awt.Color.black);
        getContentPane().add(lab3);

        int width = 90;
        int x = 120;

        now_button = new JButton("手  動");
        now_button.setBounds(x, 44, width, 24);
        now_button.setLocale(new Locale("ja","JP"));
        now_button.setFont(new java.awt.Font("dialog", 0, 12));
        now_button.setBorder(new Flush3DBorder());
        now_button.setForeground(java.awt.Color.white);
        now_button.setBackground(COLOR_MANUAL);
        now_button.addActionListener(new ResetButton());
        getContentPane().add(now_button);

        ButtonGroup mode_group = new ButtonGroup();

        mode_auto = new JRadioButton("自  動",new ImageIcon("images/rb.gif"));
        mode_auto.setBounds(x+10, 78, width, 24);
        mode_auto.setPressedIcon(new ImageIcon("images/rbp.gif"));
        mode_auto.setRolloverIcon(new ImageIcon("images/rbr.gif"));
        mode_auto.setRolloverSelectedIcon(new ImageIcon("images/rbrs.gif"));
        mode_auto.setSelectedIcon(new ImageIcon("images/rbs.gif"));
        mode_auto.setFocusPainted(false);
        mode_auto.setBorderPainted(false);
        mode_auto.setContentAreaFilled(false);
        mode_auto.setSelected(false);
        getContentPane().add(mode_auto);
        mode_group.add(mode_auto);

        mode_manual = new JRadioButton("手  動",new ImageIcon("images/rb.gif"));
        mode_manual.setBounds(x+10, 102, width, 24);
        mode_manual.setPressedIcon(new ImageIcon("images/rbp.gif"));
        mode_manual.setRolloverIcon(new ImageIcon("images/rbr.gif"));
        mode_manual.setRolloverSelectedIcon(new ImageIcon("images/rbrs.gif"));
        mode_manual.setSelectedIcon(new ImageIcon("images/rbs.gif"));
        mode_manual.setFocusPainted(false);
        mode_manual.setBorderPainted(false);
        mode_manual.setContentAreaFilled(false);
        mode_manual.setSelected(false);
        getContentPane().add(mode_manual);
        mode_group.add(mode_manual);

        mode_none = new JRadioButton("NONE");
        mode_none.setBounds(x+10, 126, width, 24);
        mode_none.setVisible(false);
        mode_none.setSelected(false);
        getContentPane().add(mode_none);
        mode_group.add(mode_none);

        updateTh = new UpdateThread();
        updateTh.setPriority(Thread.MIN_PRIORITY);
        updateTh.start();
    }

    /*******************************************************
     *
     *******************************************************/
    private boolean setSendStatus(){

        if(mode_none.isSelected()){
            send_status = MODE_NONE;
            return false;
        }
        if(mode_auto.isSelected()){
            send_status = MODE_AUTO;
        }
        if(mode_manual.isSelected()){
            send_status = MODE_MANUAL;
        }
        return true;
    }

    /*******************************************************
     *
     *******************************************************/
    public boolean setDefault(){

//@@        CZSystem.log("CZCMSControlMode","setDefault()");
        mode_none.setSelected(true);
        send_status = MODE_NONE;
        //制御モードのセット
        setData(MODE_NONE);
        return true;
    }

    /*******************************************************
     *
     *******************************************************/
    public boolean setData(int md){
        now_mode = md;

        String mode = null;
        Color  color = null;
//@@        CZSystem.log("CZCMSControlMode","setData() [" + now_mode + "]");
        switch(now_mode){
            case MODE_MANUAL :
                mode  = CZSystemDefine.PROC_MODE[MODE_MANUAL];
                color = COLOR_MANUAL;
            break;

            case MODE_AUTO :
                mode  = CZSystemDefine.PROC_MODE[MODE_AUTO];
                color = COLOR_MANUAL;
            break;

            default :
                mode  = new String("不  明");
                color = COLOR_MANUAL;
            break;
        }

        now_button.setText(mode);
        now_button.setBackground(color);
        return true;
    }

    /*******************************************************
     *
     *******************************************************/
    class SendButton implements ActionListener {
        public void actionPerformed(ActionEvent ev){

            boolean ret = setSendStatus();
//@@            CZSystem.log("CZCMSControlMode","SendButton ----->[" + send_status + "]");
            if(send_status == MODE_NONE) return;
            //Send
            if(ret){
                CZSystem.CZOperateModeExchange(now_mode,send_status);
            }   
        }
    }

    /*******************************************************
     *
     *******************************************************/
    class CancelButton implements ActionListener {
        public void actionPerformed(ActionEvent ev){
            setDefault();
            setVisible(false);
        }
    }

    /*******************************************************
     *
     *******************************************************/
    class ResetButton implements ActionListener {
        public void actionPerformed(ActionEvent ev){
            mode_none.setSelected(true);
        }
    }

    /*******************************************************
     *
     *******************************************************/
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
//@@            CZSystem.log("CZCMSControlMode","UpdateThread START");

            CZSystemQueue   que = new CZSystemQueue(10);
            CZEventAdapter  adp = new CZEventAdapter(que);
            CZEventSender.addCZEventListener(adp);
            while(true){
                try{
                    CZEventCL event = (CZEventCL)que.waitObject();
//@@                    CZSystem.log("CZCMSControlMode","RECEIVE Event ---------->");
                    if(event.getEvent() == CZEventCL.PV_RECEIVE){
//@@                        CZSystem.log("CZCMSControlMode","PV_RECEIVED");
                        setData(CZSystem.getProcMode());
                    }
                    if(event.getEvent() == CZEventCL.RO_CHANGE){
//@@                        CZSystem.log("CZCMSControlMode","RO_CHANGED");
                        setData(CZSystem.getProcMode());
                    }
//@@                    CZSystem.log("CZCMSControlMode","RECEIVE Event <----------");
                }
                catch(Exception e){

                }
            } // while end
        }
    }
}
