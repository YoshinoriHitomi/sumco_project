package cz;

import java.awt.event.WindowAdapter;
import java.awt.event.WindowEvent;
import java.util.Locale;
import java.util.Vector;

import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import javax.swing.JButton;

import java.lang.String;
import java.lang.Double;
import javax.swing.JFrame;
import javax.swing.JLabel;
import javax.swing.JTextField;
import javax.swing.plaf.metal.MetalBorders.Flush3DBorder;
import javax.swing.JOptionPane;

import czclass.CZRaidStatus;

/**********************************************************
 *
 *   �b�y���C���`�o
 * @author  (KPK Co.,Ltd.)
 * @version 1.0 (2003/04/01)
 * @version 1.2 (2006/06/20)
 * Update 2013.10.21 ����n�Q�Ƌ@�\ (@20131021)
 ***********************************************************/
public class CZMain {

    static CZEventDistributer mainEventDistributer  = null;

    static JFrame       mainWin         = null;

    static CZMainMenu   mainMenu        = null;
    static CZMainPanel  mainPanel       = null;
//    static CZMainRoNo   mainRoNo        = null;
    static CZRoSelectWin   rosel = null;
    static CZPVPanel    mainPVPanel     = null;
    static CZSetPanel   mainSetPanel    = null;
    static JLabel       mainLabel[][]   = new JLabel[4][2];

	static JButton      robutton = null;
	static JLabel       roName_lab = null;

    static CZErrorField errorField  = null; 

    //_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
    //------ �������� UpdateThread Class -----
    //_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
    static class UpdateThread extends Thread {
        UpdateThread(){

        }

        //
        //
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


        private void update(){

//@@            CZSystem.log("CZMain UpdateThread update","1");

            //BtNo�ύX
            mainLabel[0][1].setText(CZSystem.getBatch());

            //�v���Z�X���ύX
            mainLabel[1][1].setText(
                CZSystem.getProcName(CZSystem.getProcNo()));
            
            //�v���Z�X���ԕύX
            mainLabel[2][1].setText(
                CZSystem.timeFormat((long)(CZSystem.getProcTime() )));  

            //�v���Z�X���[�h
            String mode = null;
            switch(CZSystem.getProcMode()){
                case CZSystemDefine.PROC_MANUAL :   
                    mode = CZSystemDefine.PROC_MODE[CZSystemDefine.PROC_MANUAL];
                break;

                case CZSystemDefine.PROC_AUTO : 
                    mode = CZSystemDefine.PROC_MODE[CZSystemDefine.PROC_AUTO];
                break;

                default : mode = new String("�s  ��");
                break;
            }
            mainLabel[3][1].setText(mode);

//@@            CZSystem.log("CZMain UpdateThread update","2");
        }
    }

    //_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
    //------ �������� RaidErrorThread Class -----
    //_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
    static class RaidErrorThread extends Thread {

        private final int RAID1             = 0;
        private final int RAID5             = 1;

        private final int RAID_LOAD         = 1;

        private final int RAID_STAT_NONE    = 0;
        private final int RAID_STAT_OK      = 1;
        private final int RAID_STAT_NG      = 2;

        RaidErrorThread(){

        }

        //
        //
        public void run(){
            CZRaidStatus raid1_stat = null;
            CZRaidStatus raid5_stat = null;

            while(true){
//@@                CZSystem.log("CZMain RaidErrorThread","----- Check Start !! -----");
                raid1_stat = CZSystem.CZRaidGetStatus(RAID_LOAD,RAID1);
                raid5_stat = CZSystem.CZRaidGetStatus(RAID_LOAD,RAID5);

                if(null != raid1_stat){
                    if(RAID_STAT_NG == raid1_stat.getStatus()){
                        CZSystemSysMsg msg = new CZSystemSysMsg();
                        msg.no = -1;
                        msg.message = CZSystem.getDateTime() + "  [ RAID 1 : ��Q���� ]";
                        CZSystem.sysMessage(msg);
                    }
                }

                if(null != raid5_stat){
                    if(RAID_STAT_NG == raid5_stat.getStatus()){
                        CZSystemSysMsg msg = new CZSystemSysMsg();
                        msg.no = -1;
                        msg.message = CZSystem.getDateTime() + "  [RAID 5 : ��Q���� ]";
                        CZSystem.sysMessage(msg);
                    }
                }
                CZSystem.sleep(1000 * 60);  // �ꕪ���Ƀ`�F�b�N
            } // while end  
        }
    } // RaidErrorThread

    //_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
    //------ ��������@main Method -----
    //_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
    public static void main(String args[]){
//@@        CZSystem.log("CZMain main","----------> START !!");
//@@        CZSystem.log("CZMain main","[" + args.length + "][" + args + "]");

        boolean ret = CZSystem.init(CZSystemDefine.HOST_MODE,"main");

        if(!ret){
            CZSystem.exit(0,"System not Start !!"); 
        }
        //��ʂ��쐬����B
        mainWin = getMainWin();

        //Event���󂯎��Thread�𗧂��グ��
        mainEventDistributer = new CZEventDistributer();
        Thread th = new Thread(mainEventDistributer,"CZMain-mainEventDistributer");
        th.start();

        //UpdateThread�𗧂��グ��
        UpdateThread updateTh = new UpdateThread();
        updateTh.start();

        //RaidCheckThread�𗧂��グ��
        RaidErrorThread raidTh = new RaidErrorThread();
        raidTh.start();

    }

    /** @@@@@@@@
     * �N���C�A���g�o�[�W�����`�F�b�N
     */
    private static boolean VerChk(){
		
//		String ver = CZSystem.Client_ver_list.toString().trim();
		double sver = CZSystem.Client_ver_list;				//�T�[�o���Ǘ�
		
		double cver = Double.valueOf(CZSystem.VERSION);		//�N���C�A���g���Ǘ�
		
		CZSystem.log("���s�o�[�W����", "ver_" + CZSystem.VERSION);
		CZSystem.log("�ŐV�o�[�W����", "ver_" + sver);
		
//		if(ver.equals(CZSystem.VERSION.trim())){
		if(sver <= cver){
			CZSystem.log("�N���C�A���g�o�[�W����", "ver_" + sver);
		}
		else{
			Object msg[] = {"�N���C�A���g�̃o�[�W�������X�V����Ă��܂�",
							"�ŐV�o�[�W�����́Aver_" + CZSystem.Client_ver_list + "�ł�",
							"�ŐV���_�E�����[�h���܂����H"};
			int result = JOptionPane.showConfirmDialog( null,msg,"�N���C�A���g�o�[�W����",JOptionPane.YES_NO_OPTION);
			if( result == JOptionPane.YES_OPTION ){
				try{
					Runtime runtime = Runtime.getRuntime();
					/*runtime.exec( "cmd /C start C:/ClientDownload.exe");*/
					runtime.exec( "cmd /C start C:/ClientDownload/imari300/cz_download.bat");
				}
				catch (Exception ex){
				}
				return true;
			}
		}
		return false;
	}


    //_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
    // ��ʂ��쐬����Method
    //_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
    public static JFrame getMainWin(){
        JFrame win = null;

        try{
			// @@@@@@@@
			/*
            if( VerChk() == true ){
				CZSystem.exit(0,"Window Exit");
			}
			*/

            if(null == mainWin){
                win = new JFrame();
                //���Close�ŃV�X�e�����I������
                win.addWindowListener(
                    new WindowAdapter(){
                        public void windowClosing(WindowEvent e){
//@@                            CZSystem.log("CZMain getMainWin windowClosing","System.exit");
                            CZSystem.exit(0,"Window Exit");
                        }
                    }
                );

/*                win.setTitle("�ɖ��� 300mm CZ-SYSTEM MAIN [Version " + CZSystem.VERSION.trim() + "]");*/
                if(CZSystemDefine.ADMIN_RUN == CZSystem.getRunLevel()){
                     win.setTitle("�ɖ��� 300mm CZ-SYSTEM MAIN [Version 5.10]  ADMIN���[�h");
                }
                else if(CZSystemDefine.USER_RUN == CZSystem.getRunLevel()){
                     win.setTitle("�ɖ��� 300mm CZ-SYSTEM MAIN [Version 5.10]  USER���[�h");
                }
                else if(CZSystemDefine.REFERENCE_RUN == CZSystem.getRunLevel()){		// @20131021  �Q�ƃ��[�h�ǉ�
                     win.setTitle("�ɖ��� 300mm CZ-SYSTEM MAIN [Version 5.10]  REFERENCE���[�h");
                }

                win.setSize(1152,864);
//@@                win.setLocation(10,10);
                win.setLocation(0,0);

                //Menu Bar�𐶐�����B
                mainMenu     = new CZMainMenu();

                //Main�@Panel�𐶐�����B
                mainPanel = new CZMainPanel();

                // ����n�Q�Ƌ@�\    @20131021
                if(CZSystemDefine.REFERENCE_RUN == CZSystem.getRunLevel()){
                    mainPanel.setBackground(CZSystemDefine.DEFAULT_REFERENCE_BACKGROUND_COL);
                }else{
                    mainPanel.setBackground(CZSystemDefine.DEFAULT_BACKGROUND_COL);
                }

                //�F�ԑI���R���{�𐶐�����B
//                mainRoNo  = new CZMainRoNo();
//                mainPanel.add(mainRoNo,mainRoNo.getName());

                //�o�u�O���t�\��Panel�𐶐�����B
                mainPVPanel = new CZPVPanel();
                mainPanel.add(mainPVPanel,mainPVPanel.getName());
                Thread th = new Thread(mainPVPanel,"CZPVPanel-mainPVPanel");
                th.start();

                //���ڕ\��Panel�𐶐�����B
                mainSetPanel = new CZSetPanel();
                mainPanel.add(mainSetPanel,mainSetPanel.getName());
                th = new Thread(mainSetPanel,"CZPVPanel-mainSetPanel");
                th.start();
				mainSetPanel.setPanel(mainPVPanel);  // @20131030

				Vector ro = CZSystem.getRoNameList();
				String s = CZSystem.RoKetaChg((String)ro.elementAt(0));

				roName_lab = new JLabel(s,JLabel.CENTER);
				roName_lab.setBounds(20, 20, 100, 40);
				roName_lab.setLocale(new Locale("ja","JP"));
				roName_lab.setFont(new java.awt.Font("dialog", 0, 24));
				roName_lab.setBorder(new Flush3DBorder());
				roName_lab.setBackground(java.awt.Color.black);
				mainPanel.add(roName_lab);

				robutton = new JButton("��");
				robutton.setBounds(120, 20, 30, 40);
				robutton.setLocale(new Locale("ja","JP"));
				robutton.setFont(new java.awt.Font("dialog", 0, 18));  
				robutton.setBorder(new Flush3DBorder());
				robutton.setBackground(java.awt.Color.lightGray);  
				robutton.addActionListener(new RButton());
				mainPanel.add(robutton);

                //�o�b�`����\������B
                mainLabel[0][0] = new JLabel("�o�b�`No",JLabel.CENTER);
                mainLabel[0][0].setBounds(160, 25, 100, 30);
                mainLabel[0][0].setLocale(new Locale("ja","JP"));
                mainLabel[0][0].setFont(new java.awt.Font("dialog", 0, 18));
                mainLabel[0][0].setBorder(new Flush3DBorder());
                mainLabel[0][0].setForeground(java.awt.Color.black);
                mainPanel.add(mainLabel[0][0]);

                mainLabel[0][1] = new JLabel("XXXC-XXXA",JLabel.CENTER);
                mainLabel[0][1].setBounds(260, 25, 150, 30);
                mainLabel[0][1].setLocale(new Locale("ja","JP"));
                mainLabel[0][1].setFont(new java.awt.Font("dialog", 0, 18));
                mainLabel[0][1].setBorder(new Flush3DBorder());
                mainLabel[0][1].setForeground(java.awt.Color.black);
                mainPanel.add(mainLabel[0][1]);

                //�v���Z�X��\������B
                mainLabel[1][0] = new JLabel("�v���Z�X",JLabel.CENTER);
                mainLabel[1][0].setBounds(430, 25, 100, 30);
                mainLabel[1][0].setLocale(new Locale("ja","JP"));
                mainLabel[1][0].setFont(new java.awt.Font("dialog", 0, 18));
                mainLabel[1][0].setBorder(new Flush3DBorder());
                mainLabel[1][0].setForeground(java.awt.Color.black);
                mainPanel.add(mainLabel[1][0]);

                mainLabel[1][1] = new JLabel(CZSystem.getProcName(CZSystemDefine.READY),JLabel.CENTER);
                mainLabel[1][1].setBounds(530, 25, 150, 30);
                mainLabel[1][1].setLocale(new Locale("ja","JP"));
                mainLabel[1][1].setFont(new java.awt.Font("dialog", 0, 18));
                mainLabel[1][1].setBorder(new Flush3DBorder());
                mainLabel[1][1].setForeground(java.awt.Color.black);
                mainPanel.add(mainLabel[1][1]);

                //�v���Z�X���Ԃ�\������B
                mainLabel[2][0] = new JLabel("�v���Z�X����",JLabel.CENTER);
                mainLabel[2][0].setBounds(700, 25, 120, 30);
                mainLabel[2][0].setLocale(new Locale("ja","JP"));
                mainLabel[2][0].setFont(new java.awt.Font("dialog", 0, 18));
                mainLabel[2][0].setBorder(new Flush3DBorder());
                mainLabel[2][0].setForeground(java.awt.Color.black);
                mainPanel.add(mainLabel[2][0]);

                mainLabel[2][1] = new JLabel("000:00:00",JLabel.CENTER);
                mainLabel[2][1].setBounds(820, 25, 120, 30);
                mainLabel[2][1].setLocale(new Locale("ja","JP"));
                mainLabel[2][1].setFont(new java.awt.Font("dialog", 0, 18));
                mainLabel[2][1].setBorder(new Flush3DBorder());
                mainLabel[2][1].setForeground(java.awt.Color.black);
                mainPanel.add(mainLabel[2][1]);

                //���[�h��\������B
                mainLabel[3][0] = new JLabel("���[�h",JLabel.CENTER);
                mainLabel[3][0].setBounds(970, 25, 60, 30);
                mainLabel[3][0].setLocale(new Locale("ja","JP"));
                mainLabel[3][0].setFont(new java.awt.Font("dialog", 0, 18));
                mainLabel[3][0].setBorder(new Flush3DBorder());
                mainLabel[3][0].setForeground(java.awt.Color.black);
                mainPanel.add(mainLabel[3][0]);

                mainLabel[3][1] = new JLabel("�s  ��",JLabel.CENTER);
                mainLabel[3][1].setBounds(1030, 25, 100, 30);
                mainLabel[3][1].setLocale(new Locale("ja","JP"));
                mainLabel[3][1].setFont(new java.awt.Font("dialog", 0, 18));
                mainLabel[3][1].setBorder(new Flush3DBorder());
                mainLabel[3][1].setForeground(java.awt.Color.black);
                mainPanel.add(mainLabel[3][1]);

                //�G���[���b�Z�[�W��\���t�B�[���h�𐶐�����B
                errorField = new CZErrorField();    
                errorField.setBounds(20, 750, 800, 24);
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


    //_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
    //------ ��������ErroeMsg �o�̓t�B�[���hClass -----
    //_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
    static class CZErrorField extends JTextField {

        //
        // class ү���ނ��󂯎��Thread
        //
        private class MsgThread extends Thread {
            MsgThread(){
            }

            //
            //
            public void run(){

                CZSystemQueue   que = new CZSystemQueue(50);
                CZEventAdapter  adp = new CZEventAdapter(que);
                CZEventSender.addCZEventListener(adp);

                while(true){
                    try{
                        CZEventCL event = (CZEventCL)que.waitObject();

                        switch(event.getEvent()){
                            case CZEventCL.SYS_MESSAGE : setMessage(event);
                            break;
                        }
                    }
                    catch(Exception e){

                    }
                } // while end  
            }
        }

        // CZErrorField�̃R���X�g���N�^
        //
        CZErrorField(){

            super();

            setLocale(new Locale("ja","JP"));
            setFont(new java.awt.Font("dialog", 0, 16));
            setBorder(new Flush3DBorder());

            MsgThread msgTh = new MsgThread();
            msgTh.start();

        }
        //
        // method�@ү���ނ��Z�b�g����
        // @param event Event Object
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
    
    static class RButton implements ActionListener {  
        public void actionPerformed(ActionEvent e){ 
//@@            CZSystem.log("CZSetPanel","SetBtVal");  
			int X = mainWin.getX();
			int Y = mainWin.getY();
			rosel = new CZRoSelectWin(X,Y);
            rosel.setVisible(true);
        }
    }
}
