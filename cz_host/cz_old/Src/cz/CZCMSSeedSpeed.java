package cz;

import java.awt.Color;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.text.DecimalFormat;
import java.util.Locale;

import javax.swing.JButton;
import javax.swing.JDialog;
import javax.swing.JLabel;
import javax.swing.JSlider;
import javax.swing.event.ChangeEvent;
import javax.swing.event.ChangeListener;
import javax.swing.plaf.metal.MetalBorders.Flush3DBorder;

/***********************************************************
 *
 *   集中監視−シード速度変更用Window 
 * @author  (KPK Co.,Ltd.)
 * @version 1.0 (2003/04/01)
 *
 ***********************************************************/
public class CZCMSSeedSpeed extends JDialog {

    public final float    TIMES     = 0.0001f ;
    private DecimalFormat format    = new DecimalFormat("0.0000");

    private final int SEND_MAX      =  50000;
    private final int SEND_MIN      = -50000;

    public final Color COLOR_SELECT = java.awt.Color.red;
    public final Color COLOR_NONE   = java.awt.Color.blue;

    public int send_status          = 0;

    private JButton send_button     = null;
    private JButton send_undo       = null;
    private JButton cancel_button   = null;

    private JLabel  pro_label       = null;
    private JLabel  results_label   = null;
    private JButton instruct_label  = null;

    private JSlider pro_slider      = null;
    private JSlider results_slider  = null;
    private JSlider instruct_slider = null;

    private JLabel  man_label       = null;
    private JLabel  man_val         = null;

    private JButton up10000         = null;
    private JButton up1000          = null;
    private JButton up100           = null;
    private JButton up10            = null;

    private JButton down10000       = null;
    private JButton down1000        = null;
    private JButton down100         = null;
    private JButton down10          = null;

    private UpdateThread updateTh   = null;

    private final int PV_RESULT     = 17;
    private final int PV_PROFAIL    = 75;
    private final int PV_MANNUAL    = 93;


    //
    // ---------- コンストラクタ ---------------------------
    //
	@SuppressWarnings("unchecked")
    CZCMSSeedSpeed(){
        super();

        setTitle("シード速度");
        setSize(370,500);
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
        send_button.setBounds(20, 436, 100, 24);
        send_button.setLocale(new Locale("ja","JP"));
        send_button.setFont(new java.awt.Font("dialog", 0, 18));
        send_button.setBorder(new Flush3DBorder());
        send_button.setForeground(java.awt.Color.black);
        send_button.addActionListener(new SendButton());
        getContentPane().add(send_button);

        send_undo = new JButton("アンドゥ");
        send_undo.setBounds(120, 436, 100, 24);
        send_undo.setLocale(new Locale("ja","JP"));
        send_undo.setFont(new java.awt.Font("dialog", 0, 18));
        send_undo.setBorder(new Flush3DBorder());
        send_undo.setForeground(java.awt.Color.black);
        send_undo.addActionListener(new SendUndo());
        getContentPane().add(send_undo);

        cancel_button = new JButton("終  了");
        cancel_button.setBounds(240, 436, 100, 24);
        cancel_button.setLocale(new Locale("ja","JP"));
        cancel_button.setFont(new java.awt.Font("dialog", 0, 18));
        cancel_button.setBorder(new Flush3DBorder());
        cancel_button.setForeground(java.awt.Color.black);
        cancel_button.addActionListener(new CancelButton());
        getContentPane().add(cancel_button);

        man_label = new JLabel("手  介  量",JLabel.CENTER);
        man_label.setBounds(20, 392, 100, 24);
        man_label.setLocale(new Locale("ja","JP"));
        man_label.setFont(new java.awt.Font("dialog", 0, 12));
        man_label.setBorder(new Flush3DBorder());
        man_label.setForeground(java.awt.Color.black);
        getContentPane().add(man_label);

        man_val = new JLabel("0.0000 mm/min",JLabel.CENTER);
        man_val.setBounds(120, 392, 100, 24);
        man_val.setLocale(new Locale("ja","JP"));
        man_val.setFont(new java.awt.Font("dialog", 0, 12));
        man_val.setBorder(new Flush3DBorder());
        man_val.setForeground(java.awt.Color.black);
        getContentPane().add(man_val);

        ////////////////////
        JLabel lab1 = new JLabel("プロファイル",JLabel.CENTER);
        lab1.setBounds(20, 20, 100, 24);
        lab1.setLocale(new Locale("ja","JP"));
        lab1.setFont(new java.awt.Font("dialog", 0, 12));
        lab1.setBorder(new Flush3DBorder());
        lab1.setForeground(java.awt.Color.black);
        getContentPane().add(lab1);

        pro_label = new JLabel("0.0000 mm/min",JLabel.CENTER);
        pro_label.setBounds(20, 44, 100, 24);
        pro_label.setLocale(new Locale("ja","JP"));
        pro_label.setFont(new java.awt.Font("dialog", 0, 12));
        pro_label.setBorder(new Flush3DBorder());
        pro_label.setForeground(java.awt.Color.black);
        getContentPane().add(pro_label);

        pro_slider = new JSlider(JSlider.VERTICAL,0,50000,0);
        pro_slider.setBounds(20, 68, 100, 300);
        pro_slider.setBorder(new Flush3DBorder());
        pro_slider.setMajorTickSpacing(5000);
        pro_slider.setMinorTickSpacing(1000);
        pro_slider.setPaintLabels( true );
        pro_slider.setPaintTicks(true);

        JLabel label = null;
        label = new JLabel(" 5.0000", JLabel.CENTER );  
        label.setForeground(java.awt.Color.black);
        pro_slider.getLabelTable().put(new Integer(50000),label);

        label = new JLabel(" 4.5000", JLabel.CENTER );  
        label.setForeground(java.awt.Color.black);
        pro_slider.getLabelTable().put(new Integer(45000),label);

        label = new JLabel(" 4.0000", JLabel.CENTER );  
        label.setForeground(java.awt.Color.black);
        pro_slider.getLabelTable().put(new Integer(40000),label);

        label = new JLabel(" 3.5000", JLabel.CENTER );  
        label.setForeground(java.awt.Color.black);
        pro_slider.getLabelTable().put(new Integer(35000),label);

        label = new JLabel(" 3.0000", JLabel.CENTER );  
        label.setForeground(java.awt.Color.black);
        pro_slider.getLabelTable().put(new Integer(30000),label);

        label = new JLabel(" 2.5000", JLabel.CENTER );  
        label.setForeground(java.awt.Color.black);
        pro_slider.getLabelTable().put(new Integer(25000),label);

        label = new JLabel(" 2.0000", JLabel.CENTER );  
        label.setForeground(java.awt.Color.black);
        pro_slider.getLabelTable().put(new Integer(20000),label);

        label = new JLabel(" 1.5000", JLabel.CENTER );  
        label.setForeground(java.awt.Color.black);
        pro_slider.getLabelTable().put(new Integer(15000),label);

        label = new JLabel(" 1.0000", JLabel.CENTER );  
        label.setForeground(java.awt.Color.black);
        pro_slider.getLabelTable().put(new Integer(10000),label);

        label = new JLabel(" 0.5000", JLabel.CENTER );  
        label.setForeground(java.awt.Color.black);
        pro_slider.getLabelTable().put(new Integer(5000),label);

        label = new JLabel(" 0.0000", JLabel.CENTER );  
        label.setForeground(java.awt.Color.black);
        pro_slider.getLabelTable().put(new Integer(0),label);

        pro_slider.setLabelTable( pro_slider.getLabelTable() );
        pro_slider.addChangeListener(new ProEvent());
        getContentPane().add(pro_slider);

        ////////////////////
        JLabel lab2 = new JLabel("実    績",JLabel.CENTER);
        lab2.setBounds(120, 20, 100, 24);
        lab2.setLocale(new Locale("ja","JP"));
        lab2.setFont(new java.awt.Font("dialog", 0, 12));
        lab2.setBorder(new Flush3DBorder());
        lab2.setForeground(java.awt.Color.black);
        getContentPane().add(lab2);

        results_label = new JLabel("0.0000 mm/min",JLabel.CENTER);
        results_label.setBounds(120, 44, 100, 24);
        results_label.setLocale(new Locale("ja","JP"));
        results_label.setFont(new java.awt.Font("dialog", 0, 12));
        results_label.setBorder(new Flush3DBorder());
        results_label.setForeground(java.awt.Color.black);
        getContentPane().add(results_label);

        results_slider = new JSlider(JSlider.VERTICAL,0,50000,0);
        results_slider.setBounds(120, 68, 100, 300);
        results_slider.setBorder(new Flush3DBorder());
        results_slider.setMajorTickSpacing(5000);
        results_slider.setMinorTickSpacing(1000);
        results_slider.setPaintLabels( true );
        results_slider.setPaintTicks(true);

        label = null;
        label = new JLabel(" 5.0000", JLabel.CENTER );  
        label.setForeground(java.awt.Color.red);
        results_slider.getLabelTable().put(new Integer(50000),label);

        label = new JLabel(" 4.5000", JLabel.CENTER );  
        label.setForeground(java.awt.Color.red);
        results_slider.getLabelTable().put(new Integer(45000),label);

        label = new JLabel(" 4.0000", JLabel.CENTER );  
        label.setForeground(java.awt.Color.blue);
        results_slider.getLabelTable().put(new Integer(40000),label);

        label = new JLabel(" 3.5000", JLabel.CENTER );  
        label.setForeground(java.awt.Color.blue);
        results_slider.getLabelTable().put(new Integer(35000),label);

        label = new JLabel(" 3.0000", JLabel.CENTER );  
        label.setForeground(java.awt.Color.blue);
        results_slider.getLabelTable().put(new Integer(30000),label);

        label = new JLabel(" 2.5000", JLabel.CENTER );  
        label.setForeground(java.awt.Color.blue);
        results_slider.getLabelTable().put(new Integer(25000),label);

        label = new JLabel(" 2.0000", JLabel.CENTER );  
        label.setForeground(java.awt.Color.blue);
        results_slider.getLabelTable().put(new Integer(20000),label);

        label = new JLabel(" 1.5000", JLabel.CENTER );  
        label.setForeground(java.awt.Color.yellow);
        results_slider.getLabelTable().put(new Integer(15000),label);

        label = new JLabel(" 1.0000", JLabel.CENTER );  
        label.setForeground(java.awt.Color.yellow);
        results_slider.getLabelTable().put(new Integer(10000),label);

        label = new JLabel(" 0.5000", JLabel.CENTER );  
        label.setForeground(java.awt.Color.yellow);
        results_slider.getLabelTable().put(new Integer(5000),label);

        label = new JLabel(" 0.0000", JLabel.CENTER );  
        label.setForeground(java.awt.Color.black);
        results_slider.getLabelTable().put(new Integer(0),label);

        results_slider.setLabelTable( results_slider.getLabelTable() );
        results_slider.addChangeListener(new ResultsEvent());
        getContentPane().add(results_slider);

        ////////////////////
        JLabel lab3 = new JLabel("指  示  値",JLabel.CENTER);
        lab3.setBounds(240, 20, 100, 24);
        lab3.setLocale(new Locale("ja","JP"));
        lab3.setFont(new java.awt.Font("dialog", 0, 12));
        lab3.setBorder(new Flush3DBorder());
        lab3.setForeground(java.awt.Color.black);
        getContentPane().add(lab3);

        instruct_label = new JButton("0.0000 mm/min");
        instruct_label.setBounds(240, 44, 100, 24);
        instruct_label.setLocale(new Locale("ja","JP"));
        instruct_label.setFont(new java.awt.Font("dialog", 0, 12));
        instruct_label.setBorder(new Flush3DBorder());
        instruct_label.setForeground(java.awt.Color.black);
        instruct_label.addActionListener(new ZeroButton());
        getContentPane().add(instruct_label);

        instruct_slider = new JSlider(JSlider.VERTICAL,SEND_MIN,SEND_MAX,0);
        instruct_slider.setBounds(240, 68, 100, 300);
        instruct_slider.setBorder(new Flush3DBorder());
        instruct_slider.setMajorTickSpacing(5000);
        instruct_slider.setMinorTickSpacing(1000);
        instruct_slider.setPaintLabels( true );
        instruct_slider.setPaintTicks(true);

        label = null;
        label = new JLabel(" 5.0000", JLabel.CENTER );  
        label.setForeground(java.awt.Color.red);
        instruct_slider.getLabelTable().put(new Integer(50000),label);

        label = new JLabel(" 4.5000", JLabel.CENTER );  
        label.setForeground(java.awt.Color.red);
        instruct_slider.getLabelTable().put(new Integer(45000),label);

        label = new JLabel(" 4.0000", JLabel.CENTER );  
        label.setForeground(java.awt.Color.red);
        instruct_slider.getLabelTable().put(new Integer(40000),label);

        label = new JLabel(" 3.5000", JLabel.CENTER );  
        label.setForeground(java.awt.Color.red);
        instruct_slider.getLabelTable().put(new Integer(35000),label);

        label = new JLabel(" 3.0000", JLabel.CENTER );  
        label.setForeground(java.awt.Color.red);
        instruct_slider.getLabelTable().put(new Integer(30000),label);

        label = new JLabel(" 2.5000", JLabel.CENTER );  
        label.setForeground(java.awt.Color.red);
        instruct_slider.getLabelTable().put(new Integer(25000),label);

        label = new JLabel(" 2.0000", JLabel.CENTER );  
        label.setForeground(java.awt.Color.red);
        instruct_slider.getLabelTable().put(new Integer(20000),label);

        label = new JLabel(" 1.5000", JLabel.CENTER );  
        label.setForeground(java.awt.Color.red);
        instruct_slider.getLabelTable().put(new Integer(15000),label);

        label = new JLabel(" 1.0000", JLabel.CENTER );  
        label.setForeground(java.awt.Color.red);
        instruct_slider.getLabelTable().put(new Integer(10000),label);

        label = new JLabel(" 0.5000", JLabel.CENTER );  
        label.setForeground(java.awt.Color.red);
        instruct_slider.getLabelTable().put(new Integer(5000),label);

        label = new JLabel(" 0.0000", JLabel.CENTER );  
        label.setForeground(java.awt.Color.black);
        instruct_slider.getLabelTable().put(new Integer(0),label);

        label = new JLabel("-0.5000", JLabel.CENTER );  
        label.setForeground(java.awt.Color.blue);
        instruct_slider.getLabelTable().put(new Integer(-5000),label);

        label = new JLabel("-1.0000", JLabel.CENTER );  
        label.setForeground(java.awt.Color.blue);
        instruct_slider.getLabelTable().put(new Integer(-10000),label);

        label = new JLabel("-1.5000", JLabel.CENTER );  
        label.setForeground(java.awt.Color.blue);
        instruct_slider.getLabelTable().put(new Integer(-15000),label);

        label = new JLabel("-2.0000", JLabel.CENTER );  
        label.setForeground(java.awt.Color.blue);
        instruct_slider.getLabelTable().put(new Integer(-20000),label);

        label = new JLabel("-2.5000", JLabel.CENTER );  
        label.setForeground(java.awt.Color.blue);
        instruct_slider.getLabelTable().put(new Integer(-25000),label);

        label = new JLabel("-3.0000", JLabel.CENTER );  
        label.setForeground(java.awt.Color.blue);
        instruct_slider.getLabelTable().put(new Integer(-30000),label);

        label = new JLabel("-3.5000", JLabel.CENTER );  
        label.setForeground(java.awt.Color.blue);
        instruct_slider.getLabelTable().put(new Integer(-35000),label);

        label = new JLabel("-4.0000", JLabel.CENTER );  
        label.setForeground(java.awt.Color.blue);
        instruct_slider.getLabelTable().put(new Integer(-40000),label);

        label = new JLabel("-4.5000", JLabel.CENTER );  
        label.setForeground(java.awt.Color.blue);
        instruct_slider.getLabelTable().put(new Integer(-45000),label);

        label = new JLabel("-5.0000", JLabel.CENTER );  
        label.setForeground(java.awt.Color.blue);
        instruct_slider.getLabelTable().put(new Integer(-50000),label);

        instruct_slider.setLabelTable( instruct_slider.getLabelTable() );
        instruct_slider.addChangeListener(new InstructEvent());
        getContentPane().add(instruct_slider);


        /////////////////////////////////////
        up10000 = new JButton("↑");
        up10000.setBounds(240, 368, 25, 24);
        up10000.setLocale(new Locale("ja","JP"));
        up10000.setFont(new java.awt.Font("dialog", 0, 22));
        up10000.setBorder(new Flush3DBorder());
        up10000.setForeground(java.awt.Color.black);
        up10000.addActionListener(new UpButton(10000));
        getContentPane().add(up10000);

        up1000 = new JButton("↑");
        up1000.setBounds(265, 368, 25, 24);
        up1000.setLocale(new Locale("ja","JP"));
        up1000.setFont(new java.awt.Font("dialog", 0, 18));
        up1000.setBorder(new Flush3DBorder());
        up1000.setForeground(java.awt.Color.black);
        up1000.addActionListener(new UpButton(1000));
        getContentPane().add(up1000);

        up100 = new JButton("↑");
        up100.setBounds(290, 368, 25, 24);
        up100.setLocale(new Locale("ja","JP"));
        up100.setFont(new java.awt.Font("dialog", 0, 14));
        up100.setBorder(new Flush3DBorder());
        up100.setForeground(java.awt.Color.black);
        up100.addActionListener(new UpButton(100));
        getContentPane().add(up100);

        up10 = new JButton("↑");
        up10.setBounds(315, 368, 25, 24);
        up10.setLocale(new Locale("ja","JP"));
        up10.setFont(new java.awt.Font("dialog", 0, 10));
        up10.setBorder(new Flush3DBorder());
        up10.setForeground(java.awt.Color.black);
        up10.addActionListener(new UpButton(10));
        getContentPane().add(up10);

        ///////////////////////////////////////////////////////
        down10000 = new JButton("↓");
        down10000.setBounds(240, 392, 25, 24);
        down10000.setLocale(new Locale("ja","JP"));
        down10000.setFont(new java.awt.Font("dialog", 0, 22));
        down10000.setBorder(new Flush3DBorder());
        down10000.setForeground(java.awt.Color.black);
        down10000.addActionListener(new DownButton(-10000));
        getContentPane().add(down10000);

        down1000 = new JButton("↓");
        down1000.setBounds(265, 392, 25, 24);
        down1000.setLocale(new Locale("ja","JP"));
        down1000.setFont(new java.awt.Font("dialog", 0, 18));
        down1000.setBorder(new Flush3DBorder());
        down1000.setForeground(java.awt.Color.black);
        down1000.addActionListener(new DownButton(-1000));
        getContentPane().add(down1000);

        down100 = new JButton("↓");
        down100.setBounds(290, 392, 25, 24);
        down100.setLocale(new Locale("ja","JP"));
        down100.setFont(new java.awt.Font("dialog", 0, 14));
        down100.setBorder(new Flush3DBorder());
        down100.setForeground(java.awt.Color.black);
        down100.addActionListener(new DownButton(-100));
        getContentPane().add(down100);

        down10 = new JButton("↓");
        down10.setBounds(315, 392, 25, 24);
        down10.setLocale(new Locale("ja","JP"));
        down10.setFont(new java.awt.Font("dialog", 0, 10));
        down10.setBorder(new Flush3DBorder());
        down10.setForeground(java.awt.Color.black);
        down10.addActionListener(new DownButton(-10));
        getContentPane().add(down10);

        updateTh = new UpdateThread();
        updateTh.setPriority(Thread.MIN_PRIORITY);
        updateTh.start();   

    }

    //
    //
    //
    private boolean setSendStatus(){
        send_status = instruct_slider.getValue();
        return true;
    }

    //
    //
    //
    public boolean setDefault(){
        CZSystem.log("CZCMSSeedSpeed","setDefault()");
        send_status = 0;
        instruct_slider.setValue(0);
        return true;
    }


    /*******************************************************
     *
     *
     *
     *******************************************************/
    class SendButton implements ActionListener {
        public void actionPerformed(ActionEvent ev){
            setSendStatus();
            CZSystem.log("CZCMSSeedSpeed","SendButton ----->[" + send_status + "]");
            if(send_status == 0) return;
            //Send
            CZSystem.CZOperateSeedUp(send_status);
        }
    }


    /*******************************************************
     *
     *
     *
     *******************************************************/
    class SendUndo implements ActionListener {
        public void actionPerformed(ActionEvent ev){
            CZSystem.log("CZCMSSeedSpeed","SendSendUndo");
            if(send_status == 0) return;
            //Send
            CZSystem.CZOperateUndoSeedUp(send_status);
            setDefault();
        }
    }

    /*******************************************************
     *
     *
     *
     *******************************************************/
    class ZeroButton implements ActionListener {
        public void actionPerformed(ActionEvent ev){
            CZSystem.log("CZCMSSeedSpeed","ZeroSendUndo");
            setDefault();
        }
    }

    /*******************************************************
     *
     *
     *
     *******************************************************/
    class UpButton implements ActionListener {
        private int inc = 0;

        UpButton(int val){
            inc = val;
        }

        public void actionPerformed(ActionEvent ev){

            int val = instruct_slider.getValue();
            val += inc;
            if(SEND_MAX < val) return;
            instruct_slider.setValue(val);
        }
    }

    /*******************************************************
     *
     *
     *
     *******************************************************/
    class DownButton implements ActionListener {
        private int inc = 0;

        DownButton(int val){
            inc = val;
        }

        public void actionPerformed(ActionEvent ev){

            int val = instruct_slider.getValue();
            val += inc;
            if(SEND_MIN > val) return;
            instruct_slider.setValue(val);
        }
    }

    /*******************************************************
     *
     *
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
     *
     *
     *******************************************************/
    class ProcButton implements ActionListener {
        private int my_no = -1;

        public void actionPerformed(ActionEvent ev){
        
        }
    }

    /*******************************************************
     *
     *
     *
     *******************************************************/
    class ProEvent implements ChangeListener {
        public void stateChanged(ChangeEvent ev){

            int val = pro_slider.getValue();
            float data = (float)val * TIMES;
            CZSystem.log("CZCMSSeedSpeed","Proc Change[" + val + "]");
            pro_label.setText(format.format(data) + " mm/min");
        }
    }

    /*******************************************************
     *
     *
     *
     *******************************************************/
    class ResultsEvent implements ChangeListener {
        public void stateChanged(ChangeEvent ev){

            int val = results_slider.getValue();
            float data = (float)val * TIMES;
            CZSystem.log("CZCMSSeedSpeed","Results Change[" + val + "]");
            results_label.setText(format.format(data) + " mm/min");
        }
    }

    /*******************************************************
     *
     *
     *
     *******************************************************/
    class InstructEvent implements ChangeListener {
        public void stateChanged(ChangeEvent ev){

            int val = instruct_slider.getValue();
            float data = (float)val * TIMES;
            CZSystem.log("CZCMSSeedSpeed","Instruct Change[" + val + "]");
            if(0 < val) instruct_label.setForeground(java.awt.Color.red);   
            if(0 == val) instruct_label.setForeground(java.awt.Color.black);    
            if(0 > val) instruct_label.setForeground(java.awt.Color.blue);  
            instruct_label.setText(format.format(data) + " mm/min");
        }
    }


    /*******************************************************
     *
     *
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

            CZSystem.log("CZCMSSeedSpeed","UpdateThread START");

            CZSystemQueue   que = new CZSystemQueue(20);
            CZEventAdapter  adp = new CZEventAdapter(que);
            CZEventSender.addCZEventListener(adp);

            while(true){
                try{
                    CZEventCL event = (CZEventCL)que.waitObject();
                    if(event.getEvent() == CZEventCL.PV_RECEIVE){
                        setData();
                    }

                    if(event.getEvent() == CZEventCL.RO_CHANGE){
                        setData();
                    }
                }
                catch(Exception e){

                }
            } // while end
        }

        //
        //
        //
        private void setData(){
            float   prof;
            float   reslt;
            float   man;

            prof    = CZPV.getPVData(PV_PROFAIL);
            reslt   = CZPV.getPVData(PV_RESULT);
            man = CZPV.getPVData(PV_MANNUAL);
            pro_slider.setValue((int)(prof/TIMES));
            results_slider.setValue((int)(reslt/TIMES));
            man_val.setText(format.format(man) + " mm/min");

            String s = CZSystem.RoKetaChg(CZSystem.getRoName());
            setTitle(s + " : シード速度");
//            setTitle(CZSystem.getRoName() + " : シード速度");
            CZSystem.log("CZCMSSeedSpeed","" + format.format(prof) + " mm/min");
            CZSystem.log("CZCMSSeedSpeed","" + format.format(reslt) + " mm/min");
            CZSystem.log("CZCMSSeedSpeed","" + format.format(man) + " mm/min");
        }
    }
}
