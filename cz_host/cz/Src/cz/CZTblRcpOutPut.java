package cz;

import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.PrintWriter;
import java.util.Locale;
import java.util.Vector;

import javax.swing.JButton;
import javax.swing.JComboBox;
import javax.swing.JDialog;
import javax.swing.JLabel;
import javax.swing.JTextField;
import javax.swing.JOptionPane;
import javax.swing.plaf.metal.MetalBorders.Flush3DBorder;

/**
 *   ����e�[�u����r�pWindow 
 * @author  (KPK Co.,Ltd.)
 * @version 1.0 (2003/04/01)
 *
 */

public class CZTblRcpOutPut extends JDialog implements ActionListener{
	
	private GrupNo  cmdGrup = null;
	
	private JComboBox    cmbRcp1 = null;
	
	private JButton     hikaku_btn   = null;
	private JButton     cancel_button   = null;
	
	static JTextField   RoNameField;
	private JButton      robutton = null;
	
	String sHikakuHed = ",#,����,Min,Max,��,�P��,�l,�l,";
	String sLine   = new String("");
	String sDtOut  = new String("");
	
	//
	//
	//
	CZTblRcpOutPut(){
		super();
		
		setTitle("�o�̓��V�s�I��");
		
		setSize(460,230);
		setResizable(false);
		setModal(true);
		
		getContentPane().setLayout(null);
        // ����n�Q�Ƌ@�\    @20131021
        if(CZSystemDefine.REFERENCE_RUN == CZSystem.getRunLevel()){
            getContentPane().setBackground(CZSystemDefine.DEFAULT_REFERENCE_BACKGROUND_COL);
        }else{
            getContentPane().setBackground(CZSystemDefine.DEFAULT_BACKGROUND_COL);
        }
		
		
		JLabel  lab = new JLabel("�F��",JLabel.CENTER);
		lab.setBounds(20, 20, 100, 30);
		lab.setLocale(new Locale("ja","JP"));
		lab.setFont(new java.awt.Font("dialog", 0, 18));
		lab.setBorder(new Flush3DBorder());
		lab.setForeground(java.awt.Color.black);
		getContentPane().add(lab);
		
/*		Vector ro = CZSystem.getRoNameList();
		String s = CZSystem.RoKetaChg((String)ro.elementAt(0));
*/		
		RoNameField = new JTextField("");
		RoNameField.setBounds(120, 20, 100, 30);
		RoNameField.setFont(new java.awt.Font("dialog", 0, 18));
		getContentPane().add( RoNameField );
		
		robutton = new JButton("��");
		robutton.setBounds(220, 20, 30, 30);
		robutton.setLocale(new Locale("ja","JP"));
		robutton.setFont(new java.awt.Font("dialog", 0, 18));
		robutton.setBorder(new Flush3DBorder());
		robutton.setBackground(java.awt.Color.lightGray);  
		robutton.addActionListener(this);
		getContentPane().add(robutton);
		
		lab = new JLabel("�O���[�v",JLabel.CENTER);
		lab.setBounds(20, 70, 100, 30);
		lab.setLocale(new Locale("ja","JP"));
		lab.setFont(new java.awt.Font("dialog", 0, 18));
		lab.setBorder(new Flush3DBorder());
		lab.setForeground(java.awt.Color.black);
		getContentPane().add(lab);
		
		cmdGrup = new GrupNo();
		cmdGrup.setBounds(120, 70, 100, 30);
		getContentPane().add(cmdGrup);
		
		
		lab = new JLabel("���V�sNo",JLabel.CENTER);
		lab.setBounds(20, 100, 100, 30);
		lab.setLocale(new Locale("ja","JP"));
		lab.setFont(new java.awt.Font("dialog", 0, 18));
		lab.setBorder(new Flush3DBorder());
		lab.setForeground(java.awt.Color.black);
		getContentPane().add(lab);
		
		cmbRcp1 = new JComboBox();
		cmbRcp1.setBounds(120, 100, 300, 30);
		//�s�n�F�I�𒆃��V�s�i�[
		/*ro_from.setRcpDt();*/
		getContentPane().add(cmbRcp1);
		
		hikaku_btn = new JButton("�o�@��");
		hikaku_btn.setBounds(20, 150, 100, 24);
		hikaku_btn.setLocale(new Locale("ja","JP"));
		hikaku_btn.setFont(new java.awt.Font("dialog", 0, 18));
		hikaku_btn.setBorder(new Flush3DBorder());
		hikaku_btn.setForeground(java.awt.Color.black);
		hikaku_btn.addActionListener(new hikaku_btn_click());
		getContentPane().add(hikaku_btn);
		
		cancel_button = new JButton("�I  ��");
		cancel_button.setBounds(140, 150, 100, 24);
		cancel_button.setLocale(new Locale("ja","JP"));
		cancel_button.setFont(new java.awt.Font("dialog", 0, 18));
		cancel_button.setBorder(new Flush3DBorder());
		cancel_button.setForeground(java.awt.Color.black);
		cancel_button.addActionListener(new CancelButton());
		getContentPane().add(cancel_button);
	}
	
	
	//
	// ���b�Z�[�W�̕\��
	//
//	private boolean errorMsg(Object sTitle[],Object msg[]){
//		JOptionPane.showMessageDialog(null,msg,
//		sTitle,
//		JOptionPane.ERROR_MESSAGE);
//		return true;
//	}
	
//	private boolean infoMsg(Object sTitle[],Object msg[]){
//		JOptionPane.showMessageDialog(null,msg,
//		sTitle,
//		JOptionPane.INFORMATION_MESSAGE);
//		return true;
//	}
	
	
	/**
	* �A�N�V�������X�i�[
	* @param event 
	* @return none
	*/
	
	public void actionPerformed( ActionEvent event ) {
		Object source = event.getSource();
	
		CZRoSelectWin2 rosel;
		
		if( source == robutton){
			CZSystem.log("CZTbleRcpOutPut","�C�x���g�Q�b�g�I�I�I");
			
			rosel = new CZRoSelectWin2();
			rosel.setVisible(true);
			
			if(RoNameField.getText() != null){
				String roName;
				int g_no;
				CZSystemCtTitle   dtTile;
				roName = RoNameField.getText();
				
				g_no = cmdGrup.getGrupNo();
				
				CZSystem.log("setRcpDt", "�F�� : " + roName + "g_no : " + g_no );
				
				cmbRcp1.removeAllItems();
				
				//�I�����ꂽ�A�F�Ԃƃe�[�u���̏����i�[����B
				if(0 != CZSystemDefine.DISP_KETA_FLG){		/* ����200mm�p */
					StringBuffer a = new StringBuffer();
					a.append(roName);
					a.insert(0,"K");
					roName = a.toString();
				}
				
				Vector p_list = CZSystem.getCtTbRcp(roName, g_no);
				if (p_list != null)
				{
					for(int i = 0 ; p_list.size() > i ; i++){
						dtTile = (CZSystemCtTitle)p_list.elementAt(i);
						cmbRcp1.addItem( dtTile.r_no + " : " + dtTile.title.trim());
					}
				}
			}
		}
	}
	
	/*
	*
	*
	*
	*/
	class hikaku_btn_click implements ActionListener {
		 public void actionPerformed(ActionEvent ev){
			
			int tblGno = cmdGrup.getGrupNo();				//�I���O���[�v�m���擾
			String from_ro = RoNameField.getText();	//�I��F���擾
			String sBuf = null;								//�ϊ��p���o�b�t�@�[
			int rcp_from = 0;								//�I�����V�s
			
			if(0 != CZSystemDefine.DISP_KETA_FLG){		/* ����200mm�p */
				StringBuffer a = new StringBuffer();
				a.append(from_ro);
				a.insert(0,"K");
				from_ro = a.toString();
			}
			
			//�I�����V�s�m���擾
			sBuf = (String)cmbRcp1.getSelectedItem();
			if (sBuf != null)
			{
				if (sBuf.indexOf(" ") != -1)
					rcp_from = Integer.valueOf(sBuf.substring(0,sBuf.indexOf(" "))).intValue();
			}
			else
			{
				JOptionPane.showMessageDialog(null,"�o�̓��V�s�f�[�^�Ȃ�","���V�s�o��",JOptionPane.ERROR_MESSAGE);
//errorMsg("��r����", "��r���̃��V�sNo�Ȃ�");
				CZSystem.log("hikaku_btn_click","���V�s�Ȃ�");
				return;
			}
			
			//�I���O���[�v����
			if (tblGno == 6)
			{
				//��r�������{
				subT6Chk(RoNameField.getText(), from_ro, rcp_from);
			}
			else
			{
				//��r�������{
				subT1_5Chk(tblGno, RoNameField.getText(), from_ro, rcp_from);
			}
		}
	}
	
	private void subT1_5Chk(int tblGno, String fromDB_ro, String from_ro, int rcp_from){
//		int		tblGno;		�I�����ꂽ�O���[�vNo
//		String fromDB_ro;	�I��F���i��r���̂c�a���́j
//		String from_ro;		�I��F���i��r���̕\�����́j
//		int rcp_from;		�I�����V�s�i��r���j
//		String toDB_ro;		�I��F���i��r��̂c�a���́j
//		String to_ro;		�I��F���i��r��̕\�����́j
//		int rcp_to;			�I�����V�s�i��r��j
		
		Vector dataName = null;
		Vector data = null;
		CZSystemCtName dtName = null;
		CZSystemCtTb d = null;
		int iRec1;
		int iNameRec;
		int iMax1;
		int iMidasi1 = 0;
		int iChkDtRtc;
		int iRt;
		int tno;
		
		
		String sBuf = null;								//�ϊ��p���o�b�t�@�[
		
		//���̎擾
		dataName = CZSystem.ctTblAllNameRead(tblGno);
		
		//��r���̃��V�s���擾
		data = CZSystem.getCtAllTb(from_ro, tblGno, rcp_from);
		
		//���ݎ����擾
		String sNowDate = CZSystem.getDateTime("yyyy/MM/dd HH:mm:ss");
		
		//�t�@�C������
		File file = new File(CZSystem.RECIPE_OUTPUT_PATH, "���䃌�V�s�o��" + from_ro + "-" + "T" + tblGno + "-" + rcp_from + "_" + CZSystem.getDateTime("yyMMddHHmm") + ".csv");
		PrintWriter pr     = null;
		FileOutputStream s = null;
		
		try{
			String rs = CZSystem.RoKetaChg(from_ro);
			
			s = new FileOutputStream(file);
			pr = new PrintWriter(s);
			
			sLine = "����e�[�u���f�[�^�o�́i T" + tblGno + " �j,,,,,,,,," + sNowDate;
			pr.println(sLine);
			sLine = "";
			pr.println(sLine);
			sLine = "���������������o�̓��V�s��񁚁�����������";
			pr.println(sLine);
			sLine = "";
			pr.println(sLine);
			sLine = "�FNo1," + rs + ",";
			pr.println(sLine);
			sLine = "���V�sNo," + rcp_from + ",";
			pr.println(sLine);
			sLine = "";
			pr.println(sLine);
			
			sLine = "�����������������V�s���e��������������";
			pr.println(sLine);
			
			if ((null != dataName) &&(null != data)){
				iMax1 = data.size();
				iRec1 = 0;
				iNameRec = 0;
				
				for(int i = 0; ((iNameRec < dataName.size()) || (iRec1 < iMax1)); i++){
					/* �f�[�^�`�F�b�N */
					if (iNameRec < dataName.size()){
						dtName = (CZSystemCtName)dataName.elementAt(iNameRec);
					}
					
					if (iRec1 < iMax1){
						iChkDtRtc = 11;
						d = (CZSystemCtTb)data.elementAt(iRec1);
					}
					else
					{
						/* ���o���̂ݏo�� */
						iChkDtRtc = -99;
					}
					
					sDtOut = ",";	/* �f�[�^���N���A */
					iRt = -1;
					
					/* ******* ���o���`�F�b�N ******** */
					if (iNameRec >= dataName.size()){
						iRt = 2;	/* ���o���Ȃ� */
					}
					else if (iRec1 >= iMax1){
						iRt = 1;	/* ���ݒ荀�ڏo��(�f�[�^�����̍��ڏo��) */
					}
					else
					{
						iRt = chkDtName(dtName, d);
					}
					
					/********************************************/
					/**************** ���o���o�� ****************/
					/********************************************/
					if (iRt == 0)
					{
						/* ���o���o�̓`�F�b�N */
						if (iChkDtRtc != 0)
						{
							if (iMidasi1 != dtName.t_no)
							{
								sLine = "";
								pr.println(sLine);
								sLine = "�y" + dtName.t_no + " �F " + dtName.t_name.trim() + "�z";
								pr.println(sLine);
								
								iMidasi1 = dtName.t_no;
							}
							sDtOut = ",�k��,�q��,";
							pr.println(sDtOut);
						}
						iNameRec++;	/* ���ږ��̃��R�[�h�`�F���W */
					}
					else if (iRt == 1)
					{	/* ���ݒ荀�ڏo��(�f�[�^�����̍��ڏo��) */
						if (iMidasi1 != dtName.t_no)
						{
							sLine = "";
							pr.println(sLine);
							sLine = "�y" + dtName.t_no + " �F " + dtName.t_name.trim() + "�z";
							pr.println(sLine);
							
							iMidasi1 = dtName.t_no;
						}
						sDtOut = ",�k��,�q��,";
						pr.println(sDtOut);
						iNameRec++;	/* ���ږ��̃��R�[�h�`�F���W */
					}
					else if (iRt == 2)
					{	/* �Y�����o���Ȃ� */
						if (iMidasi1 != d.t_no)
						{
							sLine = "";
							pr.println(sLine);
							sLine = "�y" + d.t_no + "�z";
							pr.println(sLine);
							
							iMidasi1 = d.t_no;
						}
						sDtOut = ",�k��,�q��,";
						pr.println(sDtOut);
					}
					else
					{
						sDtOut = "����G���[,-,-,-,-,-,-,";
						iNameRec++;
						pr.println(sDtOut);
					}
					
					/********************************************/
					/**************** �f�[�^�o�� ****************/
					/********************************************/
					if ((iRt == 0) || (iRt == 2))
					{
						tno = d.t_no;
						
						for (int iLp=0; (iRec1 < iMax1);iLp++)
						{
							if (iLp >= 32760)
							{
								break;
							}
							
							if (iRec1 < iMax1)
							{
								d = (CZSystemCtTb)data.elementAt(iRec1);
							}
							
							if ((iRec1 >= iMax1) || (tno != d.t_no))
							{
								break;	//�^�[�Q�b�g�ύX
							}
							
							if (iChkDtRtc == 11)
							{	/* �f�[�^�Ⴄ�i�f�[�^�P�����Ȃ��j */
								sDtOut = ",";
								sDtOut += d.l_val + "," + d.r_val + ",";
								iRec1++;	/* �f�[�^�P���R�[�hUP */
								
								pr.println(sDtOut);
							}
						}
					}
				}	/* For End */
				JOptionPane.showMessageDialog(null,"�o�͂��������܂����B","�o�͏���",JOptionPane.INFORMATION_MESSAGE);
			}
			else
			{
				if (null == dataName)
				{
					sLine = "";
					pr.println(sLine);
					sLine = "����e�[�u���@���ڒ�`������܂���";
					pr.println(sLine);
				}
				
				if (null == data)
				{
					sLine = "";
					pr.println(sLine);
					sLine = "����e�[�u���̃f�[�^������܂���";
					pr.println(sLine);
				}
			}
		}
		catch(IOException e){
			if(null != pr) pr.close();
		}
		
		if(null != pr) pr.close();
	}
	
	//�s�P�`�s�T���o���`�F�b�N
	private int chkDtName(CZSystemCtName dtName, CZSystemCtTb dt){
		int  iRt = -1;	/* ���ږ��`�F�b�N���ʁ@0:���ꍀ�ڂ���@*/
						/*                     1:�F�f�[�^���ݒ��񂠂�i��`�݂̂���j�@*/
						/*					�@ 2:�Y�����ږ��Ȃ� */
						/*					�@ -1:����ُ� */
		
		/* ���ږ��̌��� */
		if (dtName.t_no == dt.t_no)
		{
			/* ���ꍀ�ڂ��� */
			/* ���̏o�� */
			/* ���ڃC���N�������g */
			iRt = 0;
		}
		else if (dtName.t_no < dt.t_no)
		{
			/* ���̏o�́i�f�[�^�Ȃ��j */
			/* ���ڃC���N�������g */
			iRt = 1;
		}
		else
		{
			/* ���̂Ȃ� */
			iRt = 2;
		}
		
		return(iRt);
	}
	
	 private void subT6Chk(String fromDB_ro, String from_ro, int rcp_from){
//		String fromDB_ro;	�I��F���i��r���̂c�a���́j
//		String from_ro;		�I��F���i��r���̕\�����́j
//		int rcp_from;		�I�����V�s�i��r���j
//		String toDB_ro;		�I��F���i��r��̂c�a���́j
//		String to_ro;		�I��F���i��r��̕\�����́j
//		int rcp_to;			�I�����V�s�i��r��j
		
		Vector dataName = null;
		Vector data = null;
		CZSystemCtT6AllName dtName = null;
		CZSystemCtT6Tb d = null;
		int iRec1;
		int iNameRec;
		int iMax1;
		int iMidasi1 = 0;
		int iMidasi2 = 0;
		int iChkDtRtc;
		int iRt;
		
		
		String sBuf = null;								//�ϊ��p���o�b�t�@�[
		
		//�s�U���̎擾
		dataName = CZSystem.ctT6AllNameRead();
		
		//��r���̂s�U���V�s���擾
		data = CZSystem.getCtT6AllTb(from_ro,rcp_from);
		
		//���ݎ����擾
		String sNowDate = CZSystem.getDateTime("yyyy/MM/dd HH:mm:ss");
		
		//�t�@�C������
		File file = new File(CZSystem.RECIPE_OUTPUT_PATH, "���䃌�V�s�o��" + from_ro + "-" + "T6" + "-" + rcp_from +"_" + CZSystem.getDateTime("yyMMddHHmm") + ".csv");
		PrintWriter pr     = null;
		FileOutputStream s = null;
		
		
		try{
			String rs = CZSystem.RoKetaChg(from_ro);
			
			s = new FileOutputStream(file);
			pr = new PrintWriter(s);
			
			sLine = "����e�[�u���f�[�^�o�́iT6�j,,,,,,,,," + sNowDate;
			pr.println(sLine);
			sLine = "";
			pr.println(sLine);
			sLine = "���������������o�̓��V�s��񁚁�����������";
			pr.println(sLine);
			sLine = "";
			pr.println(sLine);
			sLine = "�FNo1," + rs + ",";
			pr.println(sLine);
			sLine = "���V�sNo," + rcp_from + ",";
			pr.println(sLine);
			sLine = "";
			pr.println(sLine);
			
			sLine = "�����������������V�s���e��������������";
			pr.println(sLine);
			
			if ((null != dataName) &&(null != data))
			{
				iMax1 = data.size();
				iRec1 = 0;
				iNameRec = 0;
				for(int i = 0; ((iNameRec < dataName.size()) || (iRec1 < iMax1)); i++){ 
					/* �f�[�^�`�F�b�N */
					if (iNameRec < dataName.size())
					{
						dtName = (CZSystemCtT6AllName)dataName.elementAt(iNameRec);
					}
					
					if (iRec1 < iMax1)
					{
						d = (CZSystemCtT6Tb)data.elementAt(iRec1);
					}
					
					
					/* ******* �f�[�^��r ******** */
					if (iRec1 >= iMax1)
						iChkDtRtc = -99;
					else
						iChkDtRtc = 11;
					
					sDtOut = ",";	/* �f�[�^���N���A */
					iRt = -1;
					/* ******* ���o���`�F�b�N ******** */
					if (iNameRec >= dataName.size())
					{
						iRt = 2;	/* ���o���Ȃ� */
					}
					else if (iRec1 >= iMax1)
					{
						iRt = 1;	/* ���ݒ荀�ڏo��(�f�[�^�����̍��ڏo��) */
					}
					else
					{
						if (iChkDtRtc == 11)
						{
							iRt = chkDtNameT6(dtName, d);
						}
						else
						{
							//CZSystem.log(CZSystem.FATAL,"hikaku_btn_click chkDtRtc err sts [" + iChkDtRtc + "]");
							break;
						}
					}
					
					/********************************************/
					/**************** ���o���o�� ****************/
					/********************************************/
					if (iRt == 0)
					{
						/* ���o���o�̓`�F�b�N */
						if (iChkDtRtc != 0)
						{
							if ((iMidasi1 != dtName.k_no1) || (iMidasi2 != dtName.k_no2))
							{
								sLine = "";
								pr.println(sLine);
								sLine = "�y" + dtName.k_name1.trim() + "�z �F �y" + dtName.k_name2.trim() + "�z";
								pr.println(sLine);
								sLine = sHikakuHed;
								pr.println(sLine);
								
								iMidasi1 = dtName.k_no1;
								iMidasi2 = dtName.k_no2;
							}
							sDtOut = "," + dtName.k_no + "," + dtName.k_name.trim() + "," + dtName.k_min + "," + dtName.k_max + "," + dtName.k_keta + "," + dtName.k_unit.trim() + ",";
						}
						iNameRec++;	/* ���ږ��̃��R�[�h�`�F���W */
					}
					else if (iRt == 1)
					{	/* ���ݒ荀�ڏo��(�f�[�^�����̍��ڏo��) */
						if ((iMidasi1 != dtName.k_no1) || (iMidasi2 != dtName.k_no2))
						{
							sLine = "";
							pr.println(sLine);
							sLine = "�y" + dtName.k_name1.trim() + "�z �F �y" + dtName.k_name2.trim() + "�z";
							pr.println(sLine);
							sLine = sHikakuHed;
							pr.println(sLine);
							iMidasi1 = dtName.k_no1;
							iMidasi2 = dtName.k_no2;
						}
						sDtOut = "," + dtName.k_no + "," + dtName.k_name.trim() + "," + dtName.k_min + "," + dtName.k_max + "," + dtName.k_keta + "," + dtName.k_unit.trim() + ",-,-,";
						pr.println(sDtOut);
						iNameRec++;
					}
					else if (iRt == 2)
					{	/* �Y�����o���Ȃ� */
						if ((iMidasi1 != d.k_no1) || (iMidasi2 != d.k_no2))
						{
							sLine = "";
							pr.println(sLine);
							sLine = "�y" + d.k_no1 + "�z �F �y" + d.k_no2 + "�z";
							pr.println(sLine);
							sLine = sHikakuHed;
							pr.println(sLine);
							iMidasi1 = d.k_no1;
							iMidasi2 = d.k_no2;
						}
						sDtOut = "," + d.k_no + ",-,-,-,-,-,";
					}
					else
					{
						sDtOut = "����G���[,-,-,-,-,-,-,";
						iNameRec++;
						pr.println(sDtOut);
					}
					
					/********************************************/
					/**************** �f�[�^�o�� ****************/
					/********************************************/
					if ((iRt == 0) || (iRt == 2))
					{
						if (iChkDtRtc == 11)
						{	/* �f�[�^�Ⴄ�i�f�[�^�P�����Ȃ��j */
							sDtOut += d.k_val + ",";
							pr.println(sDtOut);
							iRec1++;	/* �f�[�^�P���R�[�hUP */
						}
					}
				}	/* for end */
				//infoMsg("�o�͏���","�o�͊���");
				JOptionPane.showMessageDialog(null,"�o�͂��������܂����B","�o�͏���",JOptionPane.INFORMATION_MESSAGE);
			}
			else
			{
				if (null == data)
				{
					sLine = "";
					pr.println(sLine);
					sLine = "T6�f�[�^�Ȃ�";
					pr.println(sLine);
				}
				
				if (null == dataName)
				{
					sLine = "";
					pr.println(sLine);
					sLine = "T6��`�f�[�^�Ȃ�";
					pr.println(sLine);
				}
			}
		}
		catch(IOException e){
			if(null != pr) pr.close();
		}
		
		if(null != pr) pr.close();
	}
	
	//�s�U���o���`�F�b�N
	private int chkDtNameT6(CZSystemCtT6AllName dtName, CZSystemCtT6Tb dt){
		int  iRt = -1;	/* ���ږ��`�F�b�N���ʁ@0:���ꍀ�ڂ���@*/
						/*                     1:�F�f�[�^���ݒ��񂠂�i��`�݂̂���j�@*/
						/*					�@ 2:�Y�����ږ��Ȃ� */
						/*					�@ -1:����ُ� */
		
		/* ���ږ��̌��� */
		if (dtName.k_no1 == dt.k_no1)
		{
			if (dtName.k_no2 == dt.k_no2)
			{
				if(dtName.k_no == dt.k_no)
				{
					/* ���ꍀ�ڂ��� */
					/* ���̏o�� */
					/* ���ڃC���N�������g */
					iRt = 0;
				}
				else if (dtName.k_no < dt.k_no)
				{
					/* ���̏o�́i�f�[�^�Ȃ��j */
					/* ���ڃC���N�������g */
					iRt = 1;
				}
				else
				{
					/* ���̂Ȃ� */
					iRt = 2;
				}
			}
			else if (dtName.k_no2 < dt.k_no2)
			{
				/* ���̏o�́i�f�[�^�Ȃ��j */
				/* ���ڃC���N�������g */
				iRt = 1;
			}
			else
			{
				/* ���̂Ȃ� */
				iRt = 2;
			}
		}
		else if (dtName.k_no1 < dt.k_no1)
		{
			/* ���̏o�́i�f�[�^�Ȃ��j */
			/* ���ڃC���N�������g */
			iRt = 1;
		}
		else
		{
			/* ���̂Ȃ� */
			iRt = 2;
		}
		
		return(iRt);
	}
	
	/*
	*
	*
	*
	*/
	class CancelButton implements ActionListener {
		public void actionPerformed(ActionEvent ev){
//			setDefault();
			setVisible(false);
		}
	}
	
	
	
	//
	// ����e�[�u���̃O���[�vNo���i�[����
	public class GrupNo extends JComboBox {
		String	sData[] = {"�s�P","�s�Q","�s�R","�s�S","�s�T","�s�U"};
		GrupNo(){
			super();
			try{
				setName("JComboBox1");
				setFont(new java.awt.Font("dialog", 0, 16));
				
				for(int i = 0 ; i < 6 ; i++){
					addItem(sData[i]);
				}
				setForeground(java.awt.Color.black);
				setBackground(java.awt.Color.lightGray);
				addActionListener(new ChgGrupNo());
			}
			catch (Exception e) {
				System.out.println("=========== System.log Exception [" + e + "]");
			}
		}
		public void setDefault(){
		}
		
		public String getGrupName() {
			/* T1����T6��Ԃ� */
			return((String)getSelectedItem());
		}
		
		public int getGrupNo() {
			/* T1����T6��1����6�ŕԂ� */
			int iNo = getSelectedIndex() + 1;
			return(iNo);
		}
		
		
		class ChgGrupNo implements ActionListener {
			public void actionPerformed(ActionEvent e){
				
				String roName1 = RoNameField.getText();
				int g_no = cmdGrup.getGrupNo();
				CZSystemCtTitle dtTile = null;
				
				cmbRcp1.removeAllItems();
				
				//�I�����ꂽ�A�F�Ԃƃe�[�u���̏����i�[����B
				if(0 != CZSystemDefine.DISP_KETA_FLG){		/* ����200mm�p */
					StringBuffer a = new StringBuffer();
					a.append(roName1);
					a.insert(0,"K");
					roName1 = a.toString();
				}
				
				Vector p_list = CZSystem.getCtTbRcp(roName1, g_no);
				if (p_list != null)
				{
					for(int i = 0 ; p_list.size() > i ; i++){
						dtTile = (CZSystemCtTitle)p_list.elementAt(i);
						cmbRcp1.addItem( dtTile.r_no + " : " + dtTile.title.trim());
					}
				}
			}
		}
	}
}
