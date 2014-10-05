package cz;

import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.text.DecimalFormat;
import java.util.Locale;

import javax.swing.JButton;
import javax.swing.JCheckBox;
import javax.swing.JDialog;
import javax.swing.JLabel;
import javax.swing.JSlider;
import javax.swing.event.ChangeEvent;
import javax.swing.event.ChangeListener;
import javax.swing.plaf.metal.MetalBorders.Flush3DBorder;

/***********************************************************
 *
 *   集中監視－ルツボ位置変更用Window 
 * @author  (KPK Co.,Ltd.)
 * @version 1.0 (2003/04/01)
 *
 ***********************************************************/
public class CZCMSCruciblePosition extends JDialog {

    public final float    TIMES     = 1.0f ;
    private DecimalFormat format    = new DecimalFormat("0.0");

    private final int SEND_MAX      =   25;
    private final int SEND_MIN      = -100;

    public int        send_status   = 0;

    private JButton   send_button   = null;
    private JButton   send_undo     = null;
    private JButton   cancel_button = null;

    private JCheckBox interlock     = null;

    private JLabel    results_label = null;
    private JButton   up_label      = null;
    private JButton   down_label    = null;

    private JSlider results_slider  = null;
    private JSlider up_slider       = null;
    private JSlider down_slider     = null;

    private UpdateThread updateTh   = null;
    private final int    PV_RESULT  = 22;

    //
    //
    //
	@SuppressWarnings("unchecked")
    CZCMSCruciblePosition(){
        super();

        setTitle("ルツボ位置");
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

        interlock = new JCheckBox("ロック解除");
        interlock.setBounds(20, 392, 100, 24);
        interlock.setLocale(new Locale("ja","JP"));
        interlock.setFont(new java.awt.Font("dialog", 0, 14));
        interlock.setBorder(new Flush3DBorder());
        interlock.setForeground(java.awt.Color.red);
        getContentPane().add(interlock);

        ////////////////////
        JLabel lab1 = new JLabel("実    績",JLabel.CENTER);
        lab1.setBounds(20, 20, 100, 24);
        lab1.setLocale(new Locale("ja","JP"));
        lab1.setFont(new java.awt.Font("dialog", 0, 12));
        lab1.setBorder(new Flush3DBorder());
        lab1.setForeground(java.awt.Color.black);
        getContentPane().add(lab1);

        results_label = new JLabel("0.0 mm",JLabel.CENTER);
        results_label.setBounds(20, 44, 100, 24);
        results_label.setLocale(new Locale("ja","JP"));
        results_label.setFont(new java.awt.Font("dialog", 0, 12));
        results_label.setBorder(new Flush3DBorder());
        results_label.setForeground(java.awt.Color.black);
        getContentPane().add(results_label);

        results_slider = new JSlider(JSlider.VERTICAL,-200,300,0);
        results_slider.setBounds(20, 68, 100, 300);
        results_slider.setBorder(new Flush3DBorder());
        results_slider.setMajorTickSpacing(50);
        results_slider.setMinorTickSpacing(10);
        results_slider.setPaintLabels( true );
        results_slider.setPaintTicks(true);

        JLabel label = null;
        label = new JLabel(" 300.0", JLabel.CENTER );   
        label.setForeground(java.awt.Color.black);
        results_slider.getLabelTable().put(new Integer(300),label);

        label = new JLabel(" 250.0", JLabel.CENTER );   
        label.setForeground(java.awt.Color.black);
        results_slider.getLabelTable().put(new Integer(250),label);

        label = new JLabel(" 200.0", JLabel.CENTER );   
        label.setForeground(java.awt.Color.black);
        results_slider.getLabelTable().put(new Integer(200),label);

        label = new JLabel(" 150.0", JLabel.CENTER );   
        label.setForeground(java.awt.Color.black);
        results_slider.getLabelTable().put(new Integer(150),label);

        label = new JLabel(" 100.0", JLabel.CENTER );   
        label.setForeground(java.awt.Color.black);
        results_slider.getLabelTable().put(new Integer(100),label);

        label = new JLabel("  50.0", JLabel.CENTER );   
        label.setForeground(java.awt.Color.black);
        results_slider.getLabelTable().put(new Integer(50),label);

        label = new JLabel("   0.0", JLabel.CENTER );   
        label.setForeground(java.awt.Color.black);
        results_slider.getLabelTable().put(new Integer(0),label);

        label = new JLabel(" -50.0", JLabel.CENTER );   
        label.setForeground(java.awt.Color.black);
        results_slider.getLabelTable().put(new Integer(-50),label);

        label = new JLabel("-100.0", JLabel.CENTER );   
        label.setForeground(java.awt.Color.black);
        results_slider.getLabelTable().put(new Integer(-100),label);

        label = new JLabel("-150.0", JLabel.CENTER );   
        label.setForeground(java.awt.Color.black);
        results_slider.getLabelTable().put(new Integer(-150),label);

        label = new JLabel("-200.0", JLabel.CENTER );   
        label.setForeground(java.awt.Color.black);
        results_slider.getLabelTable().put(new Integer(-200),label);

        results_slider.setLabelTable( results_slider.getLabelTable() );
        results_slider.addChangeListener(new ResultsEvent());
        getContentPane().add(results_slider);

        ////////////////////
        JLabel lab2 = new JLabel("上  昇  量",JLabel.CENTER);
        lab2.setBounds(140, 20, 100, 24);
        lab2.setLocale(new Locale("ja","JP"));
        lab2.setFont(new java.awt.Font("dialog", 0, 12));
        lab2.setBorder(new Flush3DBorder());
        lab2.setForeground(java.awt.Color.black);
        getContentPane().add(lab2);

        up_label = new JButton("0.0 mm");
        up_label.setBounds(140, 44, 100, 24);
        up_label.setLocale(new Locale("ja","JP"));
        up_label.setFont(new java.awt.Font("dialog", 0, 12));
        up_label.setBorder(new Flush3DBorder());
        up_label.setForeground(java.awt.Color.black);
        up_label.addActionListener(new UpZeroButton());
        getContentPane().add(up_label);

        up_slider = new JSlider(JSlider.VERTICAL,0,SEND_MAX,0);
        up_slider.setBounds(140, 68, 100, 300);
        up_slider.setBorder(new Flush3DBorder());
        up_slider.setMajorTickSpacing(5);
        up_slider.setMinorTickSpacing(1);
        up_slider.setPaintLabels( true );
        up_slider.setPaintTicks(true);

        label = null;
        label = new JLabel(" 25.0", JLabel.CENTER );    
        label.setForeground(java.awt.Color.red);
        up_slider.getLabelTable().put(new Integer(25),label);

        label = new JLabel(" 20.0", JLabel.CENTER );    
        label.setForeground(java.awt.Color.red);
        up_slider.getLabelTable().put(new Integer(20),label);

        label = new JLabel(" 15.0", JLabel.CENTER );    
        label.setForeground(java.awt.Color.red);
        up_slider.getLabelTable().put(new Integer(15),label);

        label = new JLabel(" 10.0", JLabel.CENTER );    
        label.setForeground(java.awt.Color.red);
        up_slider.getLabelTable().put(new Integer(10),label);

        label = new JLabel("  5.0", JLabel.CENTER );    
        label.setForeground(java.awt.Color.red);
        up_slider.getLabelTable().put(new Integer(5),label);

        label = new JLabel("  0.0", JLabel.CENTER );    
        label.setForeground(java.awt.Color.red);
        up_slider.getLabelTable().put(new Integer(0),label);

        up_slider.setLabelTable( up_slider.getLabelTable() );
        up_slider.addChangeListener(new UpEvent());
        getContentPane().add(up_slider);

        ////////////////////
        JLabel lab3 = new JLabel("下  降  量",JLabel.CENTER);
        lab3.setBounds(240, 20, 100, 24);
        lab3.setLocale(new Locale("ja","JP"));
        lab3.setFont(new java.awt.Font("dialog", 0, 12));
        lab3.setBorder(new Flush3DBorder());
        lab3.setForeground(java.awt.Color.black);
        getContentPane().add(lab3);

        down_label = new JButton("0.0 mm");
        down_label.setBounds(240, 44, 100, 24);
        down_label.setLocale(new Locale("ja","JP"));
        down_label.setFont(new java.awt.Font("dialog", 0, 12));
        down_label.setBorder(new Flush3DBorder());
        down_label.setForeground(java.awt.Color.black);
        down_label.addActionListener(new DownZeroButton());
        getContentPane().add(down_label);

        down_slider = new JSlider(JSlider.VERTICAL,SEND_MIN,0,0);
        down_slider.setBounds(240, 68, 100, 300);
        down_slider.setBorder(new Flush3DBorder());
        down_slider.setMajorTickSpacing(10);
        down_slider.setMinorTickSpacing(1);
        down_slider.setPaintLabels( true );
        down_slider.setPaintTicks(true);

        label = null;
        label = new JLabel("   0.0", JLabel.CENTER );   
        label.setForeground(java.awt.Color.blue);
        down_slider.getLabelTable().put(new Integer(0),label);

        label = new JLabel(" -10.0", JLabel.CENTER );   
        label.setForeground(java.awt.Color.blue);
        down_slider.getLabelTable().put(new Integer(-10),label);

        label = new JLabel(" -20.0", JLabel.CENTER );   
        label.setForeground(java.awt.Color.blue);
        down_slider.getLabelTable().put(new Integer(-20),label);

        label = new JLabel(" -30.0", JLabel.CENTER );   
        label.setForeground(java.awt.Color.blue);
        down_slider.getLabelTable().put(new Integer(-30),label);

        label = new JLabel(" -40.0", JLabel.CENTER );   
        label.setForeground(java.awt.Color.blue);
        down_slider.getLabelTable().put(new Integer(-40),label);

        label = new JLabel(" -50.0", JLabel.CENTER );   
        label.setForeground(java.awt.Color.blue);
        down_slider.getLabelTable().put(new Integer(-50),label);

        label = new JLabel(" -60.0", JLabel.CENTER );   
        label.setForeground(java.awt.Color.blue);
        down_slider.getLabelTable().put(new Integer(-60),label);

        label = new JLabel(" -70.0", JLabel.CENTER );   
        label.setForeground(java.awt.Color.blue);
        down_slider.getLabelTable().put(new Integer(-70),label);

        label = new JLabel(" -80.0", JLabel.CENTER );   
        label.setForeground(java.awt.Color.blue);
        down_slider.getLabelTable().put(new Integer(-80),label);

        label = new JLabel(" -90.0", JLabel.CENTER );   
        label.setForeground(java.awt.Color.blue);
        down_slider.getLabelTable().put(new Integer(-90),label);

        label = new JLabel("-100.0", JLabel.CENTER );   
        label.setForeground(java.awt.Color.blue);
        down_slider.getLabelTable().put(new Integer(-100),label);

        down_slider.setLabelTable( down_slider.getLabelTable() );
        down_slider.addChangeListener(new DownEvent());
        getContentPane().add(down_slider);

        updateTh = new UpdateThread();
        updateTh.setPriority(Thread.MIN_PRIORITY);
        updateTh.start();
    }

    //
    //
    //
    private boolean setSendStatus(){
        int val = 0;

        val = up_slider.getValue();
        if(0 < val){
            send_status = val;
            return true;
        }

        val = down_slider.getValue();
        if(0 > val){
            send_status = val;
            return true;
        }

        send_status = 0;
        return false;
    }


    //
    //
    //
    public boolean setDefault(){
//@@        CZSystem.log("CZCMSCruciblePosition","setDefault()");

        results_slider.setValue(0);
        up_slider.setValue(0);
        down_slider.setValue(0);

        interlock.setSelected(false);

        return true;
    }


    /*******************************************************
     *
     *
     *
     *******************************************************/
    class SendButton implements ActionListener {
        public void actionPerformed(ActionEvent ev){

            if(!setSendStatus()) return;

//@@            CZSystem.log("CZCMSCruciblePosition","SendButton ----->[" + send_status + "]");

            if(send_status == 0) return;

            //Send
            CZSystem.CZOperateRutuboPosition(send_status,interlock.isSelected());

            interlock.setSelected(false);
        }
    }


    /*******************************************************
     *
     *
     *
     *******************************************************/
    class SendUndo implements ActionListener {
        public void actionPerformed(ActionEvent ev){
//@@            CZSystem.log("CZCMSCruciblePosition","SendSendUndo");

            //Send
            CZSystem.CZOperateUndoRutuboPosition(send_status,true);

            setDefault();
        }
    }

    /*******************************************************
     *
     *
     *
     *******************************************************/
    class UpZeroButton implements ActionListener {
        public void actionPerformed(ActionEvent ev){
//@@            CZSystem.log("CZCMSCruciblePosition","UpZeroSendUndo");
            up_slider.setValue(0);
        }
    }


    /*******************************************************
     *
     *
     *
     *******************************************************/
    class DownZeroButton implements ActionListener {
        public void actionPerformed(ActionEvent ev){
//@@            CZSystem.log("CZCMSCruciblePosition","DownZeroSendUndo");
            down_slider.setValue(0);
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
    class ResultsEvent implements ChangeListener {
        public void stateChanged(ChangeEvent ev){

            int val = results_slider.getValue();
            float data = (float)val * TIMES;
//@@            CZSystem.log("CZCMSCruciblePosition","Results Change[" + val + "]");

            results_label.setText(format.format(data) + " mm");
        }
    }

    /*******************************************************
     *
     *
     *
     *******************************************************/
    class UpEvent implements ChangeListener {
        public void stateChanged(ChangeEvent ev){

            down_slider.setValue(0);

            int val = up_slider.getValue();
            float data = (float)val * TIMES;

//@@            CZSystem.log("CZCMSCruciblePosition","UpEvent Instruct Change[" + val + "]");

            if(0 < val) up_label.setForeground(java.awt.Color.red); 
            if(0 == val) up_label.setForeground(java.awt.Color.black);  

            up_label.setText(format.format(data) + " mm");
        }
    }

    /*******************************************************
     *
     *
     *
     *******************************************************/
    class DownEvent implements ChangeListener {
        public void stateChanged(ChangeEvent ev){

            up_slider.setValue(0);

            int val = down_slider.getValue();
            float data = (float)val * TIMES;

//@@            CZSystem.log("CZCMSCruciblePosition","DownEvent Instruct Change[" + val + "]");

            if(0 == val) down_label.setForeground(java.awt.Color.black);    
            if(0 > val) down_label.setForeground(java.awt.Color.blue);  

            down_label.setText(format.format(data) + " mm");
        }
    }

    /*******************************************************
     *
     *
     *
     *******************************************************/
    class UpdateThread extends Thread {

        // ---------- コンストラクタ
        //
        UpdateThread(){

        }

        //
        //
        public void run(){

//@@            CZSystem.log("CZCMSCruciblePosition","UpdateThread START");

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
            float   reslt;

            reslt   = CZPV.getPVData(PV_RESULT);

            results_slider.setValue((int)(reslt/TIMES));

            String s = CZSystem.RoKetaChg(CZSystem.getRoName());
            setTitle(s + " : ルツボ位置");
//            setTitle(CZSystem.getRoName() + " : ルツボ位置");

//@@            CZSystem.log("CZCMSCruciblePosition","CZCMSCruciblePosition " + format.format(reslt) + " mm");
        }
    }
}

