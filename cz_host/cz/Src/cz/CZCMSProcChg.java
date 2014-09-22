package cz;

import java.awt.Color;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.util.Locale;

import javax.swing.JButton;
import javax.swing.JDialog;
import javax.swing.JLabel;
import javax.swing.plaf.metal.MetalBorders.Flush3DBorder;

/*
 *  プロセス変更用Window 
 * @author  (KPK Co.,Ltd.)
 * @version 1.0 (2003/04/01)
 *
 */

public class CZCMSProcChg extends JDialog {

    public final int PROC_COUNT     = 10 ;

    public final Color COLOR_SELECT = java.awt.Color.red;
    public final Color COLOR_NONE   = java.awt.Color.blue;
    public final Color COLOR_NOW    = java.awt.Color.green;

    private int send_status = -1;
    private int now_proc    = -1;

    private JButton     send_button   = null;
    private JButton     cancel_button = null;

    private JButton     next_proc_button   = null;
    private JButton     proc_button[]      = new JButton[PROC_COUNT];

    private UpdateThread    updateTh       = null;
    
    //
    //
    //
    CZCMSProcChg(){
        super();

        setTitle("プロセス移行");
        setSize(1060,134);
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
        send_button.setBounds(210, 60, 90, 24);
        send_button.setLocale(new Locale("ja","JP"));
        send_button.setFont(new java.awt.Font("dialog", 0, 18));
        send_button.setBorder(new Flush3DBorder());
        send_button.setForeground(java.awt.Color.black);
        send_button.addActionListener(new SendButton());
        getContentPane().add(send_button);

        cancel_button = new JButton("終  了");
        cancel_button.setBounds(930, 60, 90, 24);
        cancel_button.setLocale(new Locale("ja","JP"));
        cancel_button.setFont(new java.awt.Font("dialog", 0, 18));
        cancel_button.setBorder(new Flush3DBorder());
        cancel_button.setForeground(java.awt.Color.black);
        cancel_button.addActionListener(new CancelButton());
        getContentPane().add(cancel_button);


        JLabel lab1 = new JLabel("プロセス",JLabel.CENTER);
        lab1.setBounds(20, 20, 100, 24);
        lab1.setLocale(new Locale("ja","JP"));
        lab1.setFont(new java.awt.Font("dialog", 0, 12));
        lab1.setBorder(new Flush3DBorder());
        lab1.setForeground(java.awt.Color.black);
        getContentPane().add(lab1);


        JLabel lab2 = new JLabel("移行先プロセス",JLabel.CENTER);
        lab2.setBounds(20, 60, 100, 24);
        lab2.setLocale(new Locale("ja","JP"));
        lab2.setFont(new java.awt.Font("dialog", 0, 12));
        lab2.setBorder(new Flush3DBorder());
        lab2.setForeground(java.awt.Color.black);
        getContentPane().add(lab2);

        next_proc_button = new JButton("無  し");
        next_proc_button.setBounds(120, 60, 90, 24);
        next_proc_button.setLocale(new Locale("ja","JP"));
        next_proc_button.setFont(new java.awt.Font("dialog", 0, 18));
        next_proc_button.setBorder(new Flush3DBorder());
        next_proc_button.setBackground(COLOR_NONE);
        next_proc_button.setForeground(java.awt.Color.white);
        next_proc_button.addActionListener(new ResetButton());
        getContentPane().add(next_proc_button);

        for(int i = 0 ; i < PROC_COUNT ; i++){
            createButton(i);
        }

        updateTh = new UpdateThread();
        updateTh.setPriority(Thread.MIN_PRIORITY);
        updateTh.start();

    }


    //
    //
    //
    private boolean createButton(int no){

        int off   = 120;
        int width = 90;

        proc_button[no] = new JButton(CZSystem.getProcName(no));
        proc_button[no].setBounds(off + no * width, 20, width, 24);
        proc_button[no].setLocale(new Locale("ja","JP"));
        proc_button[no].setFont(new java.awt.Font("dialog", 0, 18));
        proc_button[no].setBorder(new Flush3DBorder());
        proc_button[no].setBackground(COLOR_NONE);
        proc_button[no].setForeground(java.awt.Color.white);
        proc_button[no].addActionListener(new ProcButton(no));
        getContentPane().add(proc_button[no]);

        return true;
    }


    //
    //
    //
    public boolean setDefault(){
        CZSystem.log("CZCMSProcChg","setDefault()");

        send_status = -1;

        next_proc_button.setText("無  し");
        next_proc_button.setBackground(COLOR_NONE);
        next_proc_button.setForeground(java.awt.Color.white);
        return true;
    }

    //
    //
    //
    public boolean setData(int proc){
        now_proc = proc;

        for(int no = 0 ; no < PROC_COUNT ; no++){   
            if(now_proc == no){
                proc_button[no].setBackground(COLOR_NOW);
            }
            else {
                proc_button[no].setBackground(COLOR_NONE);
            }
        }

        return true;
    }


    /*
    *
    *
    *
    */
    class SendButton implements ActionListener {
        public void actionPerformed(ActionEvent ev){

            CZSystem.log("CZCMSProcChg","SendButton ----->[" + send_status + "]");

            if(0 > send_status) return;

            //Send
            CZSystem.CZOperateProcessExchange(now_proc,send_status);

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
    class ProcButton implements ActionListener {
        private int my_no = -1;

        //
        //
        //
        ProcButton(int _no){
            super();
            my_no = _no;
        }

        public void actionPerformed(ActionEvent ev){
            send_status = my_no;

            next_proc_button.setText(CZSystem.getProcName(my_no));
            next_proc_button.setBackground(COLOR_SELECT);
            next_proc_button.setForeground(java.awt.Color.white);
        }
    }

    /*
    *
    *
    *
    */
    class ResetButton implements ActionListener {
        public void actionPerformed(ActionEvent ev){
            setDefault();
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

            CZSystem.log("CZCMSProcChg","UpdateThread START");

            CZSystemQueue   que = new CZSystemQueue(20);
            CZEventAdapter  adp = new CZEventAdapter(que);
            CZEventSender.addCZEventListener(adp);

            while(true){
                try{
                    CZEventCL event = (CZEventCL)que.waitObject();

                    if(event.getEvent() == CZEventCL.PV_RECEIVE){
                        setData(CZSystem.getProcNo());
                    }

                    if(event.getEvent() == CZEventCL.RO_CHANGE){
                        setData(CZSystem.getProcNo());
                    }
                }
                catch(Exception e){

                }
            } // while end
        }
    }
}

