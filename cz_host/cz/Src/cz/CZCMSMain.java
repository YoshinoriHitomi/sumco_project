package cz;

import java.awt.event.WindowAdapter;
import java.awt.event.WindowEvent;
import java.util.Locale;

import javax.swing.JFrame;
import javax.swing.JLabel;
import javax.swing.JTextField;
import javax.swing.plaf.metal.MetalBorders.Flush3DBorder;

import czclass.CZRaidStatus;
/***********************************************************
 *
 *   ＣＺ集中監視メインＡＰ
 * @author  (KPK Co.,Ltd.)
 * @version 1.0 (2003/04/01)
 *
 ***********************************************************/
public class CZCMSMain {

    static CZEventDistributer mainEventDistributer  = null;

    static JFrame               mainWin     = null;

    static CZMainMenu           mainMenu    = null;
    static CZMainPanel          mainPanel   = null;
    static CZMainRoNo           mainRoNo    = null;

    static CZCMSControlPanel    mainCtlPanel    = null;

    static CZCMSPVPanel         mainPVPanel     = null;
    static CZCMSSetPanel        mainSetPanel    = null;

    static JLabel   mainLabel[][]   = new JLabel[4][2];
    static JLabel   btLabel         = null;
    static JLabel   procLabel       = null;
    static JLabel   procTmLabel     = null;
    static JLabel   procModeLabel   = null;

    static CZErrorField errorField  = null; 

    // ******************************************************
    // 更新Thread
    // ******************************************************
    static class UpdateThread extends Thread {

        // -------------- コンストラクタ -------------------
        UpdateThread(){

        }

        // --------------------------------------------------
        public void run(){
            CZSystemQueue   que = new CZSystemQueue(20);
            CZEventAdapter  adp = new CZEventAdapter(que);
            CZEventSender.addCZEventListener(adp);
            while(true){
                try{
                    CZEventCL event = (CZEventCL)que.waitObject();
                    if(event.getEvent() == CZEventCL.PV_RECEIVE){
                        update();   
                    }

                    if(event.getEvent() == CZEventCL.RO_CHANGE){
                        update();   
                    }
                }
                catch(Exception e){

                }
            } // while end 
        }
        // --------------------------------------------------
        private void update(){

            CZSystem.log("CZCMSMain UpdateThread update","1");

            //BtNo変更
            mainLabel[0][1].setText(CZSystem.getBatch());

            //プロセス名変更
            mainLabel[1][1].setText(
                    CZSystem.getProcName(CZSystem.getProcNo()));
                
            //プロセス時間変更
            mainLabel[2][1].setText(
                    CZSystem.timeFormat((long)(CZSystem.getProcTime() )));  

            //プロセスモード
            String mode = null;
            switch(CZSystem.getProcMode()){
                case CZSystemDefine.PROC_MANUAL : 
                    mode = CZSystemDefine.PROC_MODE[CZSystemDefine.PROC_MANUAL];
                    break;

                case CZSystemDefine.PROC_AUTO : 
                    mode = CZSystemDefine.PROC_MODE[CZSystemDefine.PROC_AUTO];
                    break;

                default : mode = new String("不  明");
                    break;
            }
            mainLabel[3][1].setText(mode);

            CZSystem.log("CZCMSMain UpdateThread update","2");
        }
    }


    // ******************************************************
    // RAID監視Thread
    // ******************************************************
    static class RaidErrorThread extends Thread {

        private final int RAID1 = 0;
        private final int RAID5 = 1;

        private final int RAID_LOAD   = 1;

        private final int RAID_STAT_NONE = 0;
        private final int RAID_STAT_OK   = 1;
        private final int RAID_STAT_NG   = 2;

        // -------------- コンストラクタ -------------------
        RaidErrorThread(){

        }
        // -------------------------------------------------
        public void run(){
            CZRaidStatus raid1_stat = null;
            CZRaidStatus raid5_stat = null;

            while(true){
                raid1_stat = CZSystem.CZRaidGetStatus(RAID_LOAD,RAID1);
                raid5_stat = CZSystem.CZRaidGetStatus(RAID_LOAD,RAID5);

                if(null != raid1_stat){
                    if(RAID_STAT_NG == raid1_stat.getStatus()){
                        CZSystemSysMsg msg = new CZSystemSysMsg();
                        msg.no = -1;
                        msg.message = CZSystem.getDateTime() + "  [ RAID 1 : 障害発生 ]";
                        CZSystem.sysMessage(msg);
                    }
                }

                if(null != raid5_stat){
                    if(RAID_STAT_NG == raid5_stat.getStatus()){
                        CZSystemSysMsg msg = new CZSystemSysMsg();
                        msg.no = -1;
                        msg.message = CZSystem.getDateTime() + "  [RAID 5 : 障害発生 ]";
                        CZSystem.sysMessage(msg);
                    }
                }
                CZSystem.log("CZCMSMain RaidErrorThread","Check !!");
                CZSystem.sleep(1000 * 60);  // 一分毎にチェック
            } // while end 
        }
    } // RaidErrorThread

    // ******************************************************
    // Main　Method
    // ******************************************************
    public static void main(String args[]){
        CZSystem.log("CZCMSMain main","START !!");
        CZSystem.log("CZCMSMain main","[" + args.length + "][" + args + "]");

        boolean ret = CZSystem.init(CZSystemDefine.HOST_MODE,"main");

        if(!ret){
            CZSystem.exit(0,"System not Start !!"); 
        }
        //画面を生成する。
        mainWin = getMainWin();
        //EventDistributerを起動する。
        mainEventDistributer = new CZEventDistributer();
        Thread th = new Thread(mainEventDistributer,"CZMain-mainEventDistributer");
        th.start();
        //更新Threadを起動する。
        UpdateThread updateTh = new UpdateThread();
        updateTh.start();
        //Raid監視Threadを起動する。
        RaidErrorThread raidTh = new RaidErrorThread();
        raidTh.start();
    }

    // ******************************************************
    // 画面の作成
    // ******************************************************
    public static JFrame getMainWin(){
        JFrame win = null;

        try{
            if(null == mainWin){
                win = new JFrame();
                // 画面　Close　の　ActionListener
                win.addWindowListener(
                    new WindowAdapter(){
                        public void windowClosing(WindowEvent e){
                            CZSystem.log("CZCMSMain getMainWin windowClosing","System.exit");

                            CZSystem.exit(0,"Window Exit");
                        }
                    }
                );

                win.setSize(1152,950);
                win.setLocation(0,0);
                //　メニューバー
                mainMenu     = new CZMainMenu();

                mainPanel = new CZMainPanel();
                // 他基地参照機能    @20131021
                if(CZSystemDefine.REFERENCE_RUN == CZSystem.getRunLevel()){
                    mainPanel.setBackground(CZSystemDefine.DEFAULT_REFERENCE_BACKGROUND_COL);
                }else{
                    mainPanel.setBackground(CZSystemDefine.DEFAULT_BACKGROUND_COL);
                }

                mainRoNo  = new CZMainRoNo();
                mainPanel.add(mainRoNo,mainRoNo.getName());

                mainCtlPanel = new CZCMSControlPanel();
                mainPanel.add(mainCtlPanel,mainCtlPanel.getName());

                Thread th;
                mainPVPanel = new CZCMSPVPanel();
                mainPanel.add(mainPVPanel,mainPVPanel.getName());
                th = new Thread(mainPVPanel,"CZCMSPVPanel-mainPVPanel");
                th.start();

                mainSetPanel = new CZCMSSetPanel();
                mainPanel.add(mainSetPanel,mainSetPanel.getName());
                th = new Thread(mainSetPanel,"CZCMSPVPanel-mainSetPanel");
                th.start();

                btLabel = new JLabel("バッチNo",JLabel.CENTER);
                btLabel.setBounds(160, 25, 100, 30);
                btLabel.setLocale(new Locale("ja","JP"));
                btLabel.setFont(new java.awt.Font("dialog", 0, 18));
                btLabel.setBorder(new Flush3DBorder());
                btLabel.setForeground(java.awt.Color.black);
                mainPanel.add(btLabel);

                mainLabel[0][1] = new JLabel("XXXC-XXXA",JLabel.CENTER);
                mainLabel[0][1].setBounds(260, 25, 150, 30);
                mainLabel[0][1].setLocale(new Locale("ja","JP"));
                mainLabel[0][1].setFont(new java.awt.Font("dialog", 0, 18));
                mainLabel[0][1].setBorder(new Flush3DBorder());
                mainLabel[0][1].setForeground(java.awt.Color.black);
                mainPanel.add(mainLabel[0][1]);

                procLabel = new JLabel("プロセス",JLabel.CENTER);
                procLabel.setBounds(430, 25, 100, 30);
                procLabel.setLocale(new Locale("ja","JP"));
                procLabel.setFont(new java.awt.Font("dialog", 0, 18));
                procLabel.setBorder(new Flush3DBorder());
                procLabel.setForeground(java.awt.Color.black);
                mainPanel.add(procLabel);

                mainLabel[1][1] = new JLabel(CZSystem.getProcName(CZSystemDefine.READY),JLabel.CENTER);
                mainLabel[1][1].setBounds(530, 25, 150, 30);
                mainLabel[1][1].setLocale(new Locale("ja","JP"));
                mainLabel[1][1].setFont(new java.awt.Font("dialog", 0, 18));
                mainLabel[1][1].setBorder(new Flush3DBorder());
                mainLabel[1][1].setForeground(java.awt.Color.black);
                mainPanel.add(mainLabel[1][1]);

                procTmLabel = new JLabel("プロセス時間",JLabel.CENTER);
                procTmLabel.setBounds(700, 25, 120, 30);
                procTmLabel.setLocale(new Locale("ja","JP"));
                procTmLabel.setFont(new java.awt.Font("dialog", 0, 18));
                procTmLabel.setBorder(new Flush3DBorder());
                procTmLabel.setForeground(java.awt.Color.black);
                mainPanel.add(procTmLabel);

                mainLabel[2][1] = new JLabel("000:00:00",JLabel.CENTER);
                mainLabel[2][1].setBounds(820, 25, 120, 30);
                mainLabel[2][1].setLocale(new Locale("ja","JP"));
                mainLabel[2][1].setFont(new java.awt.Font("dialog", 0, 18));
                mainLabel[2][1].setBorder(new Flush3DBorder());
                mainLabel[2][1].setForeground(java.awt.Color.black);
                mainPanel.add(mainLabel[2][1]);

                procModeLabel = new JLabel("モード",JLabel.CENTER);
                procModeLabel.setBounds(970, 25, 60, 30);
                procModeLabel.setLocale(new Locale("ja","JP"));
                procModeLabel.setFont(new java.awt.Font("dialog", 0, 18));
                procModeLabel.setBorder(new Flush3DBorder());
                procModeLabel.setForeground(java.awt.Color.black);
                mainPanel.add(procModeLabel);

                mainLabel[3][1] = new JLabel("不  明",JLabel.CENTER);
                mainLabel[3][1].setBounds(1030, 25, 100, 30);
                mainLabel[3][1].setLocale(new Locale("ja","JP"));
                mainLabel[3][1].setFont(new java.awt.Font("dialog", 0, 18));
                mainLabel[3][1].setBorder(new Flush3DBorder());
                mainLabel[3][1].setForeground(java.awt.Color.black);
                mainPanel.add(mainLabel[3][1]);
                //エラーメッセージ表示フィールドを生成する。
                errorField = new CZErrorField();    
                errorField.setBounds(20, 840, 800, 24);
                mainPanel.add(errorField);

                win.setContentPane(mainPanel);
                win.setJMenuBar(mainMenu);

                win.setVisible(true);
            }
        }
        catch (Throwable e) {
            CZSystem.handleException(e);
            win = null;
        }
        return win;
    }

    // ******************************************************
    // エラーメッセージフィールド　Class
    // ******************************************************
    static class CZErrorField extends JTextField {

        // --------------------------------------------------
        // ----- メッセージ受取るThread　Class --------------
        private class MsgThread extends Thread {
            // _/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
            // ----- MsgThreadのコンストラクタ --------------
            //
            MsgThread(){

            }
            // _/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
            public void run(){
                CZSystemQueue   que = new CZSystemQueue(50);
                CZEventAdapter  adp = new CZEventAdapter(que);
                CZEventSender.addCZEventListener(adp);
                while(true){
                    try{
                        CZEventCL event = (CZEventCL)que.waitObject();

                        switch(event.getEvent()){
                            case CZEventCL.SYS_MESSAGE :
                                setMessage(event);
                            break;
                        }
                    }
                    catch(Exception e){

                    }
                } // while end 
            }
        }
        // _/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
        //----- CZErrorField　の　コンストラクタ ------------
        //
        CZErrorField(){

            super();
            setLocale(new Locale("ja","JP"));
            setFont(new java.awt.Font("dialog", 0, 16));
            setBorder(new Flush3DBorder());
            MsgThread msgTh = new MsgThread();
            msgTh.start();
        }

        // _/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
        // ----- メッセージを更新する　Method ---------------
        // @param event エラーEvent
        //
        public void setMessage(CZEventCL event){

            CZSystemSysMsg m = (CZSystemSysMsg)event.getObject();

            switch(m.no){
                case -1 : setForeground(java.awt.Color.red);
                break;

                case  0 : setForeground(java.awt.Color.blue);
                break;

                default : setForeground(java.awt.Color.black);
                break;
            }
            setText(m.message);
        }
    }
}
