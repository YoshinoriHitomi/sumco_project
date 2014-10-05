package cz;

import java.awt.Font;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.util.Locale;

import javax.swing.JButton;
import javax.swing.JDialog;
import javax.swing.JLabel;
import javax.swing.JOptionPane;
import javax.swing.JScrollPane;
import javax.swing.JTextArea;
import javax.swing.plaf.metal.MetalBorders.Flush3DBorder;

import czclass.CZRaidStatus;

/**
 *  RAIDèÛë‘ï\é¶Window 
 * @author  (KPK Co.,Ltd.)
 * @version 1.0 (2003/04/01)
 */

public class CZRaidWatch extends JDialog {

    private CZRaidStatus raid1_stat = null;
    private CZRaidStatus raid5_stat = null;

    private final int RAID1 = 0;
    private final int RAID5 = 1;

    private final int RAID_RELOAD = 0;
    private final int RAID_LOAD   = 1;

    private final int RAID_STAT_NONE = 0;
    private final int RAID_STAT_OK   = 1;
    private final int RAID_STAT_NG   = 2;

    private JLabel  raid1_lab = null;
    private JLabel  raid5_lab = null;

    private JScrollPane raid1_sc = null;
    private JScrollPane raid5_sc = null;

    private Font        scroll_font = null;


    private JButton     send_button   = null;

    //
    //
    //
    CZRaidWatch(){
        super();

        setTitle("ÇqÇ`ÇhÇcèÛë‘");
        setSize(590,510);
        setResizable(false);
        setModal(true);

        getContentPane().setLayout(null);
        // ëºäÓínéQè∆ã@î\    @20131021
        if(CZSystemDefine.REFERENCE_RUN == CZSystem.getRunLevel()){
            getContentPane().setBackground(CZSystemDefine.DEFAULT_REFERENCE_BACKGROUND_COL);
        }else{
            getContentPane().setBackground(CZSystemDefine.DEFAULT_BACKGROUND_COL);
        }

        raid1_lab = new JLabel("RAID1",JLabel.CENTER);
        raid1_lab.setBounds(20, 20, 200, 24);
        raid1_lab.setLocale(new Locale("ja","JP"));
        raid1_lab.setFont(new java.awt.Font("dialog", 0, 16));
        raid1_lab.setBorder(new Flush3DBorder());
        raid1_lab.setForeground(java.awt.Color.black);
        getContentPane().add(raid1_lab);

        raid1_sc = new JScrollPane();
        raid1_sc.setBounds(20, 45, 550, 160);
        getContentPane().add(raid1_sc);

        raid5_lab = new JLabel("RAID5",JLabel.CENTER);
        raid5_lab.setBounds(20, 225, 200, 24);
        raid5_lab.setLocale(new Locale("ja","JP"));
        raid5_lab.setFont(new java.awt.Font("dialog", 0, 16));
        raid5_lab.setBorder(new Flush3DBorder());
        raid5_lab.setForeground(java.awt.Color.black);
        getContentPane().add(raid5_lab);

        raid5_sc = new JScrollPane();
        raid5_sc.setBounds(20, 250, 550, 180);
        getContentPane().add(raid5_sc);

        send_button = new JButton("èÛë‘éÊìæ");
        send_button.setBounds(20, 450, 100, 24);
        send_button.setLocale(new Locale("ja","JP"));
        send_button.setFont(new java.awt.Font("dialog", 0, 18));
        send_button.setBorder(new Flush3DBorder());
        send_button.setForeground(java.awt.Color.black);
        send_button.addActionListener(new SendButton());
        getContentPane().add(send_button);

        scroll_font = new java.awt.Font("dialog", 0, 12);
    }


    //
    //
    //
    public boolean setDefault(){
    
//@@        CZSystem.log("CZRaidWatch","setDefault()");
        getStatus(RAID_LOAD);

        CZSystem.log("CZRaidWatch","setDefault() Stat : " +
                    raid1_stat.getStatus() + " Log[" + raid1_stat.getLog() + "]");
        CZSystem.log("CZRaidWatch","setDefault() Stat : " +
                    raid5_stat.getStatus() + " Log[" + raid5_stat.getLog() + "]");

        writeStatus();

        return true;
    }


    //
    //
    //
    public boolean getStatus(int mode){

        raid1_stat = CZSystem.CZRaidGetStatus(mode,RAID1);
        raid5_stat = CZSystem.CZRaidGetStatus(mode,RAID5);

        return true;
    }


    //
    //
    //
    public boolean writeStatus(){

        String raid1_text = null;
        String raid1_log  = null;

        if(null == raid1_stat){
            raid1_text = " RAID1 : null";
            raid1_log  = "";
        }
        else{
            switch(raid1_stat.getStatus()){
                case RAID_STAT_NONE :
                    raid1_text = " RAID1 : ñ¢é¿ëï";
                    break;

                case RAID_STAT_OK :
                    raid1_text = " RAID1 : ê≥èÌ";
                    break;

                case RAID_STAT_NG :
                    raid1_text = " RAID1 : è·äQî≠ê∂";
                    break;
            } // switch end
            raid1_log = raid1_stat.getLog();
        }

        String raid5_text = null;
        String raid5_log  = null;

        if(null == raid5_stat){
            raid5_text = " RAID5 : null";
            raid5_log  = "";
        }
        else{
            switch(raid5_stat.getStatus()){
                case RAID_STAT_NONE :
                    raid5_text = " RAID5 : ñ¢é¿ëï";
                    break;

                case RAID_STAT_OK :
                    raid5_text = " RAID5 : ê≥èÌ";
                    break;

                case RAID_STAT_NG :
                    raid5_text = " RAID5 : è·äQî≠ê∂";
                    break;
            } // switch end
            raid5_log = raid5_stat.getLog();
        }

        raid1_lab.setText(raid1_text);
        raid5_lab.setText(raid5_text);

        JTextArea r1 = new JTextArea(raid1_log);
        r1.setFont(scroll_font);

        JTextArea r5 = new JTextArea(raid5_log);
        r5.setFont(scroll_font);

        raid1_sc.setViewportView(r1);
        raid5_sc.setViewportView(r5);
        return true;
    }


    //
    // ÉÅÉbÉZÅ[ÉWÇÃï\é¶
    //
    private boolean errorMsg(Object msg[]){
        JOptionPane.showMessageDialog(null,msg,
        "RAIDèÛë‘ï\é¶",
        JOptionPane.ERROR_MESSAGE);
        return true;
    }


    /*
    *
    *
    *
    */
    class SendButton implements ActionListener {
        public void actionPerformed(ActionEvent ev){
            getStatus(RAID_RELOAD);
//@@            CZSystem.log("CZRaidWatch","SendButton Stat : " +
//@@                        raid1_stat.getStatus() + " Log[" + raid1_stat.getLog() + "]");
//@@            CZSystem.log("CZRaidWatch","SendButton Stat : " +
//@@                        raid5_stat.getStatus() + " Log[" + raid5_stat.getLog() + "]");
            writeStatus();
            return ;
        }
    }
} // CZRaidWatch
