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
import javax.swing.JOptionPane;
import javax.swing.plaf.metal.MetalBorders.Flush3DBorder;

/**
 *   ����e�[�u����r�pWindow 
 * @author  (KPK Co.,Ltd.)
 * @version 1.0 (2003/04/01)
 * @Update 2013/10/17 ���V�s�ԍ��ێ��@�\ (@20131017)
 */

public class CZTblHikaku2 extends JDialog {

    private RoNo    ro_from = null;
    private RoNo    ro_to = null;

    private GrupNo  cmdGrup = null;

    private JComboBox    cmbRcp1 = null;
    private JComboBox    cmbRcp2 = null;

    private JButton     hikaku_btn   = null;
    private JButton     cancel_button   = null;

	String sHikakuHed = ",#,����,Min,Max,��,�P��,�l,�l,";
    String sLine   = new String("");
    String sDtOut  = new String("");

    // ��r��
    String sRcp1Info = new String("");  //@20131017
    int bufRcp1No = 0;  //@20131017

    // ��r��
    String sRcp2Info = new String("");  //@20131017
    int bufRcp2No = 0;  //@20131017

    //
    //
    //
    CZTblHikaku2(){
        super();

   	    setTitle("����e�[�u����r");
	
		setSize(460,320);
        setResizable(false);
        setModal(true);

        getContentPane().setLayout(null);
        // ����n�Q�Ƌ@�\    @20131021
        if(CZSystemDefine.REFERENCE_RUN == CZSystem.getRunLevel()){
            getContentPane().setBackground(CZSystemDefine.DEFAULT_REFERENCE_BACKGROUND_COL);
        }else{
            getContentPane().setBackground(CZSystemDefine.DEFAULT_BACKGROUND_COL);
        }


        JLabel  lab = new JLabel("�O���[�v",JLabel.CENTER);
        lab.setBounds(20, 20, 100, 30);
        lab.setLocale(new Locale("ja","JP"));
        lab.setFont(new java.awt.Font("dialog", 0, 18));
        lab.setBorder(new Flush3DBorder());
        lab.setForeground(java.awt.Color.black);
        getContentPane().add(lab);

        cmdGrup = new GrupNo();
        cmdGrup.setBounds(120, 20, 100, 30);
        getContentPane().add(cmdGrup);


        lab = new JLabel("��r��",JLabel.CENTER);
        lab.setBounds(20, 70, 100, 30);
        lab.setLocale(new Locale("ja","JP"));
        lab.setFont(new java.awt.Font("dialog", 0, 18));
        lab.setBorder(new Flush3DBorder());
        lab.setForeground(java.awt.Color.black);
        getContentPane().add(lab);

        ro_from = new RoNo(1);
        ro_from.setBounds(120, 70, 100, 30);
        getContentPane().add(ro_from);


        lab = new JLabel("���V�sNo",JLabel.CENTER);
        lab.setBounds(20, 100, 100, 30);
        lab.setLocale(new Locale("ja","JP"));
        lab.setFont(new java.awt.Font("dialog", 0, 18));
        lab.setBorder(new Flush3DBorder());
        lab.setForeground(java.awt.Color.black);
        getContentPane().add(lab);

        cmbRcp1 = new JComboBox();
        cmbRcp1.setBounds(120, 100, 300, 30);
		cmbRcp1.setFocusable(false);	/* 2007.08.22 */

		//�s�n�F�I�𒆃��V�s�i�[
		ro_from.setRcpDt();
        getContentPane().add(cmbRcp1);

        lab = new JLabel("��r��",JLabel.CENTER);
        lab.setBounds(20, 140, 100, 30);
        lab.setLocale(new Locale("ja","JP"));
        lab.setFont(new java.awt.Font("dialog", 0, 18));
        lab.setBorder(new Flush3DBorder());
        lab.setForeground(java.awt.Color.black);
        getContentPane().add(lab);

        ro_to = new RoNo(2);
        ro_to.setBounds(120, 140, 100, 30);
        getContentPane().add(ro_to);


        lab = new JLabel("���V�sNo",JLabel.CENTER);
        lab.setBounds(20, 170, 100, 30);
        lab.setLocale(new Locale("ja","JP"));
        lab.setFont(new java.awt.Font("dialog", 0, 18));
        lab.setBorder(new Flush3DBorder());
        lab.setForeground(java.awt.Color.black);
        getContentPane().add(lab);

        cmbRcp2 = new JComboBox();
        cmbRcp2.setBounds(120, 170, 300, 30);
		cmbRcp2.setFocusable(false);	/* 2007.08.22 */
		//�s�n�F�I�𒆃��V�s�i�[
		ro_to.setRcpDt();
        getContentPane().add(cmbRcp2);


        hikaku_btn = new JButton("��@�r");
        hikaku_btn.setBounds(20, 240, 100, 24);
        hikaku_btn.setLocale(new Locale("ja","JP"));
        hikaku_btn.setFont(new java.awt.Font("dialog", 0, 18));
        hikaku_btn.setBorder(new Flush3DBorder());
        hikaku_btn.setForeground(java.awt.Color.black);
        hikaku_btn.addActionListener(new hikaku_btn_click());
        getContentPane().add(hikaku_btn);

        cancel_button = new JButton("�I  ��");
        cancel_button.setBounds(140, 240, 100, 24);
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
//    private boolean errorMsg(Object sTitle[],Object msg[]){
//        JOptionPane.showMessageDialog(null,msg,
//        sTitle,
//        JOptionPane.ERROR_MESSAGE);
//        return true;
//    }

//    private boolean infoMsg(Object sTitle[],Object msg[]){
//        JOptionPane.showMessageDialog(null,msg,
//        sTitle,
//        JOptionPane.INFORMATION_MESSAGE);
//        return true;
//    }



    /*
    *
    *
    *
    */
    class hikaku_btn_click implements ActionListener {
        public void actionPerformed(ActionEvent ev){
			
			int tblGno = cmdGrup.getGrupNo();				//�I���O���[�v�m���擾
			String from_ro = ro_from.GetSelectedRoban();	//�I��F���i��r���j�@�擾
			String to_ro = ro_to.GetSelectedRoban();		//�I��F���i��r��j�@�擾
			String sBuf = null;								//�ϊ��p���o�b�t�@�[
			int rcp_from = 0;								//�I�����V�s�i��r���j�@
			int rcp_to = 0;									//�I�����V�s�i��r��j�@


			//�I�����V�s�m���i��r���j�擾
			sBuf = (String)cmbRcp1.getSelectedItem();
			if (sBuf != null)
			{
				if (sBuf.indexOf(" ") != -1)
					rcp_from = Integer.valueOf(sBuf.substring(0,sBuf.indexOf(" "))).intValue();
			}
			else
			{
                JOptionPane.showMessageDialog(null,"��r���̃��V�sNo�Ȃ�","��r����",JOptionPane.ERROR_MESSAGE);
//errorMsg("��r����", "��r���̃��V�sNo�Ȃ�");
				CZSystem.log("hikaku_btn_click","���V�s�Ȃ��i�P�j");
				return;
			}

			//�I�����V�s�m���i��r��j�擾
			sBuf = (String)cmbRcp2.getSelectedItem();
			if (sBuf != null)
			{
				if (sBuf.indexOf(" ") != -1)
					rcp_to = Integer.valueOf(sBuf.substring(0,sBuf.indexOf(" "))).intValue();
			}
			else
			{
//                errorMsg("��r����","��r��̃��V�sNo�Ȃ�");
                JOptionPane.showMessageDialog(null,"��r��̃��V�sNo�Ȃ�","��r����",JOptionPane.ERROR_MESSAGE);
				CZSystem.log("hikaku_btn_click","���V�s�Ȃ��i�Q�j");
				return;
			}
			

			CZSystem.log("hikaku_btn_click","from_ro=[" + from_ro + "] to_ro[" + to_ro + "]");
			
			//�I���O���[�v����
			if (tblGno == 6)
			{
				//��r�������{
				subT6Chk(ro_from.GetSelectedRobanDBname(), from_ro, rcp_from, 
						 ro_to.GetSelectedRobanDBname(),   to_ro,   rcp_to);
			}
			else
			{
				//��r�������{
				subT1_5Chk(tblGno, ro_from.GetSelectedRobanDBname(), from_ro, rcp_from, 
						 ro_to.GetSelectedRobanDBname(),   to_ro,   rcp_to);
			}
		}
   }

    private void subT1_5Chk(int tblGno, String fromDB_ro, String from_ro, int rcp_from, String toDB_ro, String to_ro, int rcp_to){
//		int		tblGno;		�I�����ꂽ�O���[�vNo
//		String fromDB_ro;	�I��F���i��r���̂c�a���́j
//		String from_ro;		�I��F���i��r���̕\�����́j
//		int rcp_from;		�I�����V�s�i��r���j
//		String toDB_ro;		�I��F���i��r��̂c�a���́j
//		String to_ro;		�I��F���i��r��̕\�����́j
//		int rcp_to;			�I�����V�s�i��r��j
	
        Vector dataName = null;
        Vector data = null;
        Vector data2 = null;
		CZSystemCtName dtName = null;
		CZSystemCtTb d = null;
		CZSystemCtTb d2 = null;
		int iRec1;
		int iRec2;
		int iNameRec;
		int iMax1;
		int iMax2;
		int iMidasi1 = 0;
		int iMidasi2 = 0;
		int iChkDtRtc;
		int iRt;
		int tno;


		String sBuf = null;								//�ϊ��p���o�b�t�@�[
		
		//���̎擾
		dataName = CZSystem.ctTblAllNameRead(tblGno);

		//��r���̃��V�s���擾
		data = CZSystem.getCtAllTb(fromDB_ro, tblGno, rcp_from);

		//��r��̂s�U���V�s���擾
		data2 = CZSystem.getCtAllTb(toDB_ro, tblGno, rcp_to);

		//���ݎ����擾
		String sNowDate = CZSystem.getDateTime("yyyy/MM/dd HH:mm:ss");

		//�t�@�C������
        File file = new File(CZSystem.SOGYO_OUTPUT_PATH, "�����r" + from_ro + "_"  + to_ro + "_"  + CZSystem.getDateTime("yyMMddHHmm") + ".csv");
        PrintWriter pr     = null;
		FileOutputStream s = null;

        try{
			s = new FileOutputStream(file);
			pr = new PrintWriter(s);

			sLine = "����e�[�u����r�i T" + tblGno + " �j,,,,,,,,," + sNowDate;
			pr.println(sLine);
			sLine = "";
			pr.println(sLine);
			sLine = "����������������r������������������";
			pr.println(sLine);
			sLine = "";
			pr.println(sLine);
			sLine = "�FNo1," + from_ro + ",";
			pr.println(sLine);
			sLine = "���V�sNo," + rcp_from + ",";
			pr.println(sLine);
			sLine = "";
			pr.println(sLine);

			sLine = "�FNo2," + to_ro + ",";
			pr.println(sLine);
			sLine = "���V�sNo," + rcp_to + ",";
			pr.println(sLine);
			sLine = "";
			pr.println(sLine);
			sLine = "����������������r���ʁ�������������";
			pr.println(sLine);
		
	        if ((null != dataName) &&(null != data) && (null != data2))
			{
	            iMax1 = data.size();
	            iMax2 = data2.size();
				iRec1 = 0;
				iRec2 = 0;
				iNameRec = 0;
	            for(int i = 0; ((iNameRec < dataName.size()) || (iRec1 < iMax1) || (iRec2 < iMax2)); i++){ 
					/* �f�[�^�`�F�b�N */
					if (iNameRec < dataName.size())
					{
						dtName = (CZSystemCtName)dataName.elementAt(iNameRec);
					}

					if (iRec1 < iMax1)
					{
						d = (CZSystemCtTb)data.elementAt(iRec1);
					}

					if (iRec2 < iMax2)
					{
						d2 = (CZSystemCtTb)data2.elementAt(iRec2);
					}

					/* ******* �f�[�^��r ******** */
					if ((iRec1 >= iMax1) && (iRec2 >= iMax2))
				    {
						/* ���o���̂ݏo�� */
						iChkDtRtc = -99;
					}
					else if (iRec1 >= iMax1)
				    {
						/* data2 nomi */
						iChkDtRtc = 12;
					}
					else if (iRec2 >= iMax2)
					{
						/* data1 nomi */
				    	iChkDtRtc = 11;
					}
					else
				    {
						CZSystemCtTb dt = null;
						CZSystemCtTb dt2 = null;
						int iRecChk1;
						int iRecChk2;

										//�S���R�[�h����or���o���Ȃ�

						iRecChk1 = iRec1;
						iRecChk2 = iRec2;
						if (iRecChk1 < iMax1)
						{
							dt = (CZSystemCtTb)data.elementAt(iRecChk1);
						}

						if (iRecChk2 < iMax2)
						{
							dt2 = (CZSystemCtTb)data2.elementAt(iRecChk2);
						}


						/* ���ږ��̌��� */
						if (dt.t_no == dt2.t_no)
						{
							tno = dt.t_no;
							iChkDtRtc = 0;
			CZSystem.log("�f�[�^�`�F�b�NST","[" + iRecChk1 + "] [" + iMax1 + "]" + "[" + iRecChk2 + "] [" + iMax2 + "]");
							for (int iLp=0; ((iRecChk1 < iMax1) || (iRecChk2 < iMax2));iLp++)
							{
								if (iRecChk1 < iMax1)
								{
									dt = (CZSystemCtTb)data.elementAt(iRecChk1);
								}

								if (iRecChk2 < iMax2)
								{
									dt2 = (CZSystemCtTb)data2.elementAt(iRecChk2);
								}

								//�l�̃`�F�b�N 2008.01.28
								if (((tno != dt.t_no) && (tno != dt2.t_no)) ||
								    ((iRecChk1 >= iMax1) && (tno != dt2.t_no)) ||
								    ((tno != dt.t_no) && (iRecChk2 >= iMax2)))
								{	//�����Ⴄ�i�`�F�b�N�I���j
									break;
								}

								if (((tno != dt.t_no) && (tno == dt2.t_no)) ||
								    ((tno == dt.t_no) && (tno != dt2.t_no)))
								{
									/* �f�[�^�Ⴄ */
									iChkDtRtc = 1;
									break;
								}

								if ((tno == dt.t_no) && (tno == dt2.t_no))
								{
									/* ���ꍀ�ڂ��� */
									if ((dt.k_no == dt2.k_no) && 
									    (dt.l_val == dt2.l_val) && 
									    (dt.r_val == dt2.r_val))
									{
										/* �f�[�^���� */
									}
									else
									{
										/* �f�[�^�Ⴄ */
										iChkDtRtc = 1;
										break;
									}
								}
								iRecChk1++;
								iRecChk2++;
			CZSystem.log("�f�[�^�`�F�b�N��","[" + iRecChk1 + "] [" + iMax1 + "]" + "[" + iRecChk2 + "] [" + iMax2 + "]");
							}
			CZSystem.log("�f�[�^�`�F�b�N�I��","[" + iRecChk1 + "] [" + iMax1 + "]" + "[" + iRecChk2 + "] [" + iMax2 + "]");

						}
						else if (dt.t_no < dt2.t_no)
						{
							/* �f�[�^�P�f�[�^���� */
							iChkDtRtc = 11;
						}
						else
						{
							/* �f�[�^�Q�f�[�^���� */
							iChkDtRtc = 12;
						}
					}

					sDtOut = ",";	/* �f�[�^���N���A */
					iRt = -1;
					/* ******* ���o���`�F�b�N ******** */
					if (iNameRec >= dataName.size())
				    {
						iRt = 2;	/* ���o���Ȃ� */
					}
					else if ((iRec1 >= iMax1) && (iRec2 >= iMax2))
				    {
						iRt = 1;	/* ���ݒ荀�ڏo��(�f�[�^�����̍��ڏo��) */
					}
					else
					{
						if ((iChkDtRtc == 0) ||
						    (iChkDtRtc == 1) ||
						    (iChkDtRtc == 11))
						{
					    	iRt = chkDtName(dtName, d);
						}
						else if (iChkDtRtc == 12)
						{
							/* �f�[�^�Ⴄ�i�f�[�^�Q�����Ȃ��j */
					    	iRt = chkDtName(dtName, d2);
					    }
						else
						{
							CZSystem.log("hikaku_btn_click","chkDtRtc err sts [" + iChkDtRtc + "]");
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
							if (iMidasi1 != dtName.t_no)
							{
								sLine = "";
								pr.println(sLine);
								sLine = "�y" + dtName.t_no + " �F " + dtName.t_name.trim() + "�z";
								pr.println(sLine);

								iMidasi1 = dtName.t_no;
							}
							sDtOut = ",�k��,�q��,,�k��,�q��,";
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
						sDtOut = ",�k��,�q��,,�k��,�q��,";
						pr.println(sDtOut);
						iNameRec++;	/* ���ږ��̃��R�[�h�`�F���W */
					}
					else if (iRt == 2)
					{	/* �Y�����o���Ȃ� */
						if (iChkDtRtc == 12)
						{
							if (iMidasi1 != d2.t_no)
							{
								sLine = "";
								pr.println(sLine);
								sLine = "�y" + d2.t_no + "�z";
								pr.println(sLine);

								iMidasi1 = d2.t_no;
							}
							sDtOut = ",�k��,�q��,,�k��,�q��,";
							pr.println(sDtOut);
						}
						else
						{
							if (iMidasi1 != d.t_no)
							{
								sLine = "";
								pr.println(sLine);
								sLine = "�y" + d.t_no + "�z";
								pr.println(sLine);

								iMidasi1 = d.t_no;
							}
							sDtOut = ",�k��,�q��,,�k��,�q��,";
							pr.println(sDtOut);
						}
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
						if (iChkDtRtc == 12)
							tno = d2.t_no;

			CZSystem.log("�f�[�^�o�͊J�n","[" + iRec1 + "] [" + iMax1 + "]" + "[" + iRec2 + "] [" + iMax2 + "]");
						for (int iLp=0; ((iRec1 < iMax1) || (iRec2 < iMax2));iLp++)
						{
			CZSystem.log("�f�[�^�o��--start","[" + iChkDtRtc + "] [" + iRt + "]");
							if (iLp >= 32760)
							{
			CZSystem.log("�f�[�^�o��--�����I��","Loop�@OVER");
								break;
							}

							if (iRec1 < iMax1)
							{
								d = (CZSystemCtTb)data.elementAt(iRec1);
							}

							if (iRec2 < iMax2)
							{
								d2 = (CZSystemCtTb)data2.elementAt(iRec2);
							}

							if (((iRec1 >= iMax1) || (tno != d.t_no)) && 
								((iRec2 >= iMax2) || (tno != d2.t_no)))
							{
								break;	//�^�[�Q�b�g�ύX
							}

							if (iChkDtRtc == 0)
							{	/* �f�[�^���� */
								/* ��O���� */
								/* ��`�ɂȂ����ڂ̏ꍇ�́A�f�[�^�������ł� */
								/* �l��\�� */
								if (iRt == 2) 
								{
									sDtOut = ",";
									if ((iRec1 < iMax1) && (tno == d.t_no))
									{
										sDtOut += d.l_val + "," + d.r_val + ",";
										iRec1++;	/* �f�[�^�P���R�[�hUP */
									}
									else
										sDtOut += "-,-,";

									if ((iRec2 < iMax2) && (tno == d2.t_no))
									{
										sDtOut += "," + d2.l_val + "," + d2.r_val + ",";
										iRec2++;	/* �f�[�^�Q���R�[�hUP */
									}
									else
										sDtOut += ",-,-,";

									pr.println(sDtOut);
								}
								else
								{
									iRec1++;	/* �f�[�^�P���R�[�hUP */
									iRec2++;	/* �f�[�^�Q���R�[�hUP */
								}


							}
							else if (iChkDtRtc == 1)
							{	/* �f�[�^�Ⴄ�i�l���Ⴄ�j */
								//�l�̃`�F�b�N
								sDtOut = ",";
								if ((iRec1 < iMax1) && (tno == d.t_no))
								{
									sDtOut += d.l_val + "," + d.r_val + ",";
									iRec1++;	/* �f�[�^�P���R�[�hUP */
								}
								else
									sDtOut += "-,-,";

								if ((iRec2 < iMax2) && (tno == d2.t_no))
								{
									sDtOut += "," + d2.l_val + "," + d2.r_val + ",";
									iRec2++;	/* �f�[�^�Q���R�[�hUP */
								}
								else
									sDtOut += ",-,-,";

								pr.println(sDtOut);
							}
							else if (iChkDtRtc == 11)
							{	/* �f�[�^�Ⴄ�i�f�[�^�P�����Ȃ��j */
								sDtOut = ",";
								if ((iRec1 < iMax1) && (tno == d.t_no))
								{
									sDtOut += d.l_val + "," + d.r_val + ",";
									iRec1++;	/* �f�[�^�P���R�[�hUP */
								}
								else
									sDtOut += "-,-,";

								sDtOut += ",-,-,";

								pr.println(sDtOut);
							}
							else if (iChkDtRtc == 12)
							{	/* �f�[�^�Ⴄ�i�f�[�^�Q�����Ȃ��j */
								sDtOut = ",";
								sDtOut += "-,-,";

								if ((iRec2 < iMax2) && (tno == d2.t_no))
								{
									sDtOut += "," + d2.l_val + "," + d2.r_val + ",";
									iRec2++;	/* �f�[�^�Q���R�[�hUP */
								}
								else
									sDtOut += ",-,-,";

								pr.println(sDtOut);
						    }
						}
			CZSystem.log("�f�[�^�o�͒�","[" + iRec1 + "] [" + iMax1 + "]" + "[" + iRec2 + "] [" + iMax2 + "]");
					}
				}	/* for end */
			CZSystem.log("�f�[�^�o�͊���","[" + iRec1 + "] [" + iMax1 + "]" + "[" + iRec2 + "] [" + iMax2 + "]");
//                infoMsg("��r����", "��r����");
                JOptionPane.showMessageDialog(null,"��r���������܂����B","��r����",JOptionPane.INFORMATION_MESSAGE);
	        }
			else
			{
		        if ((null == data) && (null != data2))
				{
					CZSystem.log("hikaku_btn_click","��r�P�f�[�^�Ȃ�");
					sLine = "";
					pr.println(sLine);
					sLine = "��r�P�f�[�^�Ȃ�";
					pr.println(sLine);
//	                errorMsg("��r����","��r���̏�񂪂���܂���");
	                JOptionPane.showMessageDialog(null,"��r���̏�񂪂���܂���","��r����",JOptionPane.ERROR_MESSAGE);
				}
				else
		        if ((null != data) && (null == data2))
				{
					CZSystem.log("hikaku_btn_click","��r�Q�f�[�^�Ȃ�");
					sLine = "";
					pr.println(sLine);
					sLine = "��r�Q�f�[�^�Ȃ�";
					pr.println(sLine);
//	                errorMsg("��r����","��r��̏�񂪂���܂���");
	                JOptionPane.showMessageDialog(null,"��r��̏�񂪂���܂���","��r����",JOptionPane.ERROR_MESSAGE);
				}
				else
				{
					CZSystem.log("hikaku_btn_click","��r�P�A�Q�f�[�^�Ȃ�");
					sLine = "";
					pr.println(sLine);
					sLine = "��r�P�A�Q�f�[�^�Ȃ�";
					pr.println(sLine);
//	                errorMsg("��r����","��r��񂪂���܂���");
	                JOptionPane.showMessageDialog(null,"��r��񂪂���܂���","��r����",JOptionPane.ERROR_MESSAGE);
				}
			}
        }
        catch(IOException e){
            if(null != pr) pr.close();
        }

        if(null != pr) pr.close();
	}

    private void subT6Chk(String fromDB_ro, String from_ro, int rcp_from, String toDB_ro, String to_ro, int rcp_to){
//		String fromDB_ro;	�I��F���i��r���̂c�a���́j
//		String from_ro;		�I��F���i��r���̕\�����́j
//		int rcp_from;		�I�����V�s�i��r���j
//		String toDB_ro;		�I��F���i��r��̂c�a���́j
//		String to_ro;		�I��F���i��r��̕\�����́j
//		int rcp_to;			�I�����V�s�i��r��j
	
        Vector dataName = null;
        Vector data = null;
        Vector data2 = null;
		CZSystemCtT6AllName dtName = null;
		CZSystemCtT6Tb d = null;
		CZSystemCtT6Tb d2 = null;
		int iRec1;
		int iRec2;
		int iNameRec;
		int iMax1;
		int iMax2;
		int iMidasi1 = 0;
		int iMidasi2 = 0;
		int iChkDtRtc;
		int iRt;


		String sBuf = null;								//�ϊ��p���o�b�t�@�[
		
		//�s�U���̎擾
		dataName = CZSystem.ctT6AllNameRead();

		//��r���̂s�U���V�s���擾
		data = CZSystem.getCtT6AllTb(fromDB_ro,rcp_from);

		//��r��̂s�U���V�s���擾
		data2 = CZSystem.getCtT6AllTb(toDB_ro,rcp_to);

		//���ݎ����擾
		String sNowDate = CZSystem.getDateTime("yyyy/MM/dd HH:mm:ss");

		//�t�@�C������
        File file = new File(CZSystem.SOGYO_OUTPUT_PATH, "�����r" + from_ro + "_"  + to_ro + "_"  + CZSystem.getDateTime("yyMMddHHmm") + ".csv");
        PrintWriter pr     = null;
		FileOutputStream s = null;

        try{
			s = new FileOutputStream(file);
			pr = new PrintWriter(s);

			sLine = "����e�[�u����r�i�s�U�j,,,,,,,,," + sNowDate;
			pr.println(sLine);
			sLine = "";
			pr.println(sLine);
			sLine = "����������������r������������������";
			pr.println(sLine);
			sLine = "";
			pr.println(sLine);
			sLine = "�FNo1," + from_ro + ",";
			pr.println(sLine);
			sLine = "���V�sNo," + rcp_from + ",";
			pr.println(sLine);
			sLine = "";
			pr.println(sLine);

			sLine = "�FNo2," + to_ro + ",";
			pr.println(sLine);
			sLine = "���V�sNo," + rcp_to + ",";
			pr.println(sLine);
			sLine = "";
			pr.println(sLine);
			sLine = "����������������r���ʁ�������������";
			pr.println(sLine);
		
	        if ((null != dataName) &&(null != data) && (null != data2))
			{
	            iMax1 = data.size();
	            iMax2 = data2.size();
				iRec1 = 0;
				iRec2 = 0;
				iNameRec = 0;
	            for(int i = 0; ((iNameRec < dataName.size()) || (iRec1 < iMax1) || (iRec2 < iMax2)); i++){ 
					/* �f�[�^�`�F�b�N */
					if (iNameRec < dataName.size())
					{
						dtName = (CZSystemCtT6AllName)dataName.elementAt(iNameRec);
					}

					if (iRec1 < iMax1)
					{
						d = (CZSystemCtT6Tb)data.elementAt(iRec1);
					}

					if (iRec2 < iMax2)
					{
						d2 = (CZSystemCtT6Tb)data2.elementAt(iRec2);
					}

					/* ******* �f�[�^��r ******** */
					if ((iRec1 >= iMax1) && (iRec2 >= iMax2))
				    	iChkDtRtc = -99;
					else if (iRec1 >= iMax1)
				    	iChkDtRtc = 12;
					else if (iRec2 >= iMax2)
				    	iChkDtRtc = 11;
					else
				    	iChkDtRtc = chkDtT6(d, d2);

					sDtOut = ",";	/* �f�[�^���N���A */
					iRt = -1;
					/* ******* ���o���`�F�b�N ******** */
					if (iNameRec >= dataName.size())
				    {
						iRt = 2;	/* ���o���Ȃ� */
					}
					else if ((iRec1 >= iMax1) && (iRec2 >= iMax2))
				    {
						iRt = 1;	/* ���ݒ荀�ڏo��(�f�[�^�����̍��ڏo��) */
					}
					else
					{
						if ((iChkDtRtc == 0) ||
						    (iChkDtRtc == 1) ||
						    (iChkDtRtc == 11))
						{
					    	iRt = chkDtNameT6(dtName, d);
						}
						else if (iChkDtRtc == 12)
						{
							/* �f�[�^�Ⴄ�i�f�[�^�Q�����Ȃ��j */
					    	iRt = chkDtNameT6(dtName, d2);
					    }
						else
						{
							CZSystem.log("hikaku_btn_click","chkDtRtc err sts [" + iChkDtRtc + "]");
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
						if (iChkDtRtc == 12)
						{
							if ((iMidasi1 != d2.k_no1) || (iMidasi2 != d2.k_no2))
							{
								sLine = "";
								pr.println(sLine);
								sLine = "�y" + d2.k_no1 + "�z �F �y" + d2.k_no2 + "�z";
								pr.println(sLine);
								sLine = sHikakuHed;
								pr.println(sLine);
								iMidasi1 = d2.k_no1;
								iMidasi2 = d2.k_no2;
							}
							sDtOut = "," + d2.k_no + ",-,-,-,-,-,";
						}
						else
						{
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
						if (iChkDtRtc == 0)
						{	/* �f�[�^���� */
							/* ��O���� */
							/* ��`�ɂȂ����ڂ̏ꍇ�́A�f�[�^�������ł� */
							/* �l��\�� */
							if (iRt == 2) 
							{
								sDtOut += d.k_val + "," + d2.k_val;
								pr.println(sDtOut);
							}
							iRec1++;	/* �f�[�^�P���R�[�hUP */
							iRec2++;	/* �f�[�^�Q���R�[�hUP */
						}
						else if (iChkDtRtc == 1)
						{	/* �f�[�^�Ⴄ�i�l���Ⴄ�j */
							sDtOut += d.k_val + "," + d2.k_val;
							pr.println(sDtOut);
							iRec1++;	/* �f�[�^�P���R�[�hUP */
							iRec2++;	/* �f�[�^�Q���R�[�hUP */
						}
						else if (iChkDtRtc == 11)
						{	/* �f�[�^�Ⴄ�i�f�[�^�P�����Ȃ��j */
							sDtOut += d.k_val + ",-";
							pr.println(sDtOut);
							iRec1++;	/* �f�[�^�P���R�[�hUP */
						}
						else if (iChkDtRtc == 12)
						{	/* �f�[�^�Ⴄ�i�f�[�^�Q�����Ȃ��j */
							sDtOut += "-," + d2.k_val;
							pr.println(sDtOut);
							iRec2++;	/* �f�[�^�Q���R�[�hUP */
					    }
					}
				}	/* for end */
//                infoMsg("��r����","��r����");
                JOptionPane.showMessageDialog(null,"��r���������܂����B","��r����",JOptionPane.INFORMATION_MESSAGE);
	        }
			else
			{
		        if ((null == data) && (null != data2))
				{
					CZSystem.log("hikaku_btn_click","��r�P�f�[�^�Ȃ�");
					sLine = "";
					pr.println(sLine);
					sLine = "��r�P�f�[�^�Ȃ�";
					pr.println(sLine);
//	                errorMsg("��r����","��r���̏�񂪂���܂���");
	                JOptionPane.showMessageDialog(null,"��r���̏�񂪂���܂���","��r����",JOptionPane.ERROR_MESSAGE);
				}
				else
		        if ((null != data) && (null == data2))
				{
					CZSystem.log("hikaku_btn_click","��r�Q�f�[�^�Ȃ�");
					sLine = "";
					pr.println(sLine);
					sLine = "��r�Q�f�[�^�Ȃ�";
					pr.println(sLine);
//	                errorMsg("��r����","��r��̏�񂪂���܂���");
	                JOptionPane.showMessageDialog(null,"��r��̏�񂪂���܂���","��r����",JOptionPane.ERROR_MESSAGE);
				}
				else
				{
					CZSystem.log("hikaku_btn_click","��r�P�A�Q�f�[�^�Ȃ�");
					sLine = "";
					pr.println(sLine);
					sLine = "��r�P�A�Q�f�[�^�Ȃ�";
					pr.println(sLine);
//	                errorMsg("��r����","��r��񂪂���܂���");
	                JOptionPane.showMessageDialog(null,"��r��񂪂���܂���","��r����",JOptionPane.ERROR_MESSAGE);
				}
			}
        }
        catch(IOException e){
            if(null != pr) pr.close();
        }

        if(null != pr) pr.close();
	}



	//�s�U�f�[�^��r
    private int chkDtT6(CZSystemCtT6Tb dt,CZSystemCtT6Tb dt2){
		int  iRt = -1;	/* ���ږ��`�F�b�N���ʁ@0:�����l�@*/
						/*                     1:�Ⴄ�l�@*/
						/*					�@ 11:�f�[�^�P�̂݃f�[�^���� */
						/*					�@ 12:�f�[�^�Q�̂݃f�[�^���� */
						/*					�@ -1:����ُ� */

		/* ���ږ��̌��� */
		if (dt.k_no1 == dt2.k_no1)
		{
			if (dt.k_no2 == dt2.k_no2)
			{
				if(dt.k_no == dt2.k_no)
				{
					/* ���ꍀ�ڂ��� */
					if (dt.k_val == dt2.k_val)
					{
						/* �f�[�^���� */
						iRt = 0;
					}
					else
					{
						/* �f�[�^�Ⴄ */
						iRt = 1;
					}
				}
				else if (dt.k_no < dt2.k_no)
				{
					/* �f�[�^�P�f�[�^���� */
					iRt = 11;
				}
				else
				{
					/* �f�[�^�Q�f�[�^���� */
					iRt = 12;
				}
			}
			else if (dt.k_no2 < dt2.k_no2)
			{
				/* �f�[�^�P�f�[�^���� */
				iRt = 11;
			}
			else
			{
				/* �f�[�^�Q�f�[�^���� */
				iRt = 12;
			}
		}
		else if (dt.k_no1 < dt2.k_no1)
		{
			/* �f�[�^�P�f�[�^���� */
			iRt = 11;
		}
		else
		{
			/* �f�[�^�Q�f�[�^���� */
			iRt = 12;
		}

		return(iRt);
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
//            setDefault();
            setVisible(false);
        }
    }


	//
	// �F�Ԃ��i�[����
    public class RoNo extends JComboBox {
		int	iMode = 0;		/* 1:From 2:To */

        RoNo(int iModePt){
            super();
            try{
				iMode = iModePt;
                setName("JComboBox1");
                setFont(new java.awt.Font("dialog", 0, 16));

                Vector ro = CZSystem.getRoNameList();

                if(null == ro){
                    System.out.println("Not Ro No1");
                }

                for(int i = 0 ; ro.size() > i ; i++){
					String s = CZSystem.RoKetaChg((String)ro.elementAt(i));	// 20050725 �F�F�\�������ύX
                    addItem(s);
                }
                setForeground(java.awt.Color.black);
                setBackground(java.awt.Color.lightGray);
				setFocusable(false);	/* 2007.08.22 */
                addActionListener(new ChgRoNo());
            }
            catch (Exception e) {
				System.out.println("=========== System.log Exception [" + e + "]");
            }
        }
        public void setDefault(){
        }

		public String GetSelectedRoban() {
			Object os = getSelectedItem();
			return(os.toString());
		}

		public String GetSelectedRobanDBname() {
			String ro;

			if( 0 != CZSystemDefine.DISP_KETA_FLG){
				StringBuffer a = new StringBuffer();
				a.append((String)getSelectedItem());
				a.insert(0,"K");
				String s = a.toString();
				ro = s;
			} else {
				ro = (String)getSelectedItem();
			}
			return ro;
		}


        class ChgRoNo implements ActionListener {
            public void actionPerformed(ActionEvent e){
				setRcpDt();
			}
		}

		//���ݑI�𒆂̘F�Ԃ̃��V�s���i�[
		public void setRcpDt() {
			String roName;
			int g_no;
			int index1 = -1;  // @20131017
			int index2 = -1;  // @20131017
			CZSystemCtTitle   dtTile;
			roName = GetSelectedRobanDBname();

			g_no = cmdGrup.getGrupNo();

			CZSystem.log("setRcpDt", "�F�� : " + roName + "g_no : " + g_no + "���[�h[" + iMode + "]");

			if (iMode == 1){
				//�I�����V�s�m���i��r���j�擾 @20131017
				sRcp1Info = (String)cmbRcp1.getSelectedItem();
				if (sRcp1Info != null)
				{
					if (sRcp1Info.indexOf(" ") != -1)
						bufRcp1No = Integer.valueOf(sRcp1Info.substring(0,sRcp1Info.indexOf(" "))).intValue();
				}
	            CZSystem.log("CZTblHikaku2 ChgRcpNo","�i��r���j���V�s�ԍ��ύX");
	            CZSystem.log("CZTblHikaku2 ChgRcpNo","�i��r���j���V�s�ԍ��擾�F"+ bufRcp1No);
				// @20131017
			}else if (iMode == 2){
				//�I�����V�s�m���i��r��j�擾 @20131017
				sRcp2Info = (String)cmbRcp2.getSelectedItem();
				if (sRcp2Info != null)
				{
					if (sRcp2Info.indexOf(" ") != -1)
						bufRcp2No = Integer.valueOf(sRcp2Info.substring(0,sRcp2Info.indexOf(" "))).intValue();
				}
	            CZSystem.log("CZTblHikaku2 ChgRcp2No","�i��r��j���V�s�ԍ��ύX");
	            CZSystem.log("CZTblHikaku2 ChgRcp2No","�i��r��j���V�s�ԍ��擾�F"+ bufRcp2No);
				// @20131017
			}

			if (iMode == 1)			//from
				cmbRcp1.removeAllItems();
			else if (iMode == 2)	//to
				cmbRcp2.removeAllItems();
			
			//�I�����ꂽ�A�F�Ԃƃe�[�u���̏����i�[����B
			
			Vector p_list = CZSystem.getCtTbRcp(roName, g_no);
			if (p_list != null)
			{
				for(int i = 0 ; p_list.size() > i ; i++){
					dtTile = (CZSystemCtTitle)p_list.elementAt(i);
					if (iMode == 1)			//from
					{
						cmbRcp1.addItem( dtTile.r_no + " : " + dtTile.title.trim());
						
						// @20131017 �O��I���������V�sNo���ύX�F�Ԑ�ɓ������V�sNo�����邩�H
						int icnt = cmbRcp1.getItemCount();
						CZSystem.log("CZTblHikaku2 setRcpDt","�R���{�{�b�N�X�� ���V�s���F" + icnt);
						CZSystem.log("CZTblHikaku2 setRcpDt","�R���{�{�b�N�X�� ���V�sNo�F" + dtTile.r_no);
						CZSystem.log("CZTblHikaku2 setRcpDt","�R���{�{�b�N�X�� �O��I�����V�sNo�F" + bufRcp1No);
						if (dtTile.r_no == bufRcp1No)  // �������V�sNo������ꍇ��index�擾
						{
							index1 = icnt;
							CZSystem.log("CZTblHikaku2 setRcpDt","�R���{�{�b�N�X Index1���F" + index1);
						}
						// @20131017
					}
					else if (iMode == 2)	//to
					{
						cmbRcp2.addItem( dtTile.r_no + " : " +  dtTile.title.trim());
						
						// @20131017 �O��I���������V�sNo���ύX�F�Ԍ��ɓ������V�sNo�����邩�H
						int icnt = cmbRcp2.getItemCount();
						CZSystem.log("CZTblHikaku2 setRcpDt","�R���{�{�b�N�X�� ���V�s���F" + icnt);
						CZSystem.log("CZTblHikaku2 setRcpDt","�R���{�{�b�N�X�� ���V�sNo�F" + dtTile.r_no);
						CZSystem.log("CZTblHikaku2 setRcpDt","�R���{�{�b�N�X�� �O��I�����V�sNo�F" + bufRcp2No);
						if (dtTile.r_no == bufRcp2No)  // �������V�sNo������ꍇ��index�擾
						{
							index2 = icnt;
							CZSystem.log("CZTblHikaku2 setRcpDt","�R���{�{�b�N�X Index2���F" + index2);
						}
						// @20131017
					}
				}
				//@20131017 �O��I�����V�s�ԍ����w��
				CZSystem.log("CZTblHikaku2 setRcpDt","�O��I�����V�s�ԍ�index���w�� : " + index1);
				if (index1 != -1)
				{
					cmbRcp1.setSelectedIndex(index1-1);
				}
				// @20131017
				//@20131017 �O��I�����V�s�ԍ����w��
				CZSystem.log("CZTblHikaku2 setRcpDt","�O��I�����V�s�ԍ�index2���w�� : " + index2);
				if (index2 != -1)
				{
					cmbRcp2.setSelectedIndex(index2-1);
				}
				// @20131017
			}
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
				setFocusable(false);	/* 2007.08.22 */
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
				
				String roName1 = ro_from.GetSelectedRobanDBname();
				String roName2 = ro_to.GetSelectedRobanDBname();
				int g_no = cmdGrup.getGrupNo();
				int index1 = -1;  // @20131017
				int index2 = -1;  // @20131017
				CZSystemCtTitle dtTile = null;

				CZSystem.log("ChgRoNo", "�F�ԂP : " + roName1 + "�F�ԂQ�F" + roName2 + "g_no : " + g_no);
				
				//�I�����V�s�m���i��r���j�擾 @20131017
				sRcp1Info = (String)cmbRcp1.getSelectedItem();
				if (sRcp1Info != null)
				{
					if (sRcp1Info.indexOf(" ") != -1)
						bufRcp1No = Integer.valueOf(sRcp1Info.substring(0,sRcp1Info.indexOf(" "))).intValue();
				}
	            CZSystem.log("CZTblHikaku2 ChgRcp1No","�i��r���j���V�s�ԍ��ύX");
	            CZSystem.log("CZTblHikaku2 ChgRcp1No","�i��r���j���V�s�ԍ��擾�F"+ bufRcp1No);
				// @20131017

				//�I�����V�s�m���i��r��j�擾 @20131017
				sRcp2Info = (String)cmbRcp2.getSelectedItem();
				if (sRcp2Info != null)
				{
					if (sRcp2Info.indexOf(" ") != -1)
						bufRcp2No = Integer.valueOf(sRcp2Info.substring(0,sRcp2Info.indexOf(" "))).intValue();
				}
	            CZSystem.log("CZTblHikaku2 ChgRcp2No","�i��r��j���V�s�ԍ��ύX");
	            CZSystem.log("CZTblHikaku2 ChgRcp2No","�i��r��j���V�s�ԍ��擾�F"+ bufRcp2No);
				// @20131017

				cmbRcp1.removeAllItems();
				cmbRcp2.removeAllItems();
				
				//�I�����ꂽ�A�F�Ԃƃe�[�u���̏����i�[����B
				Vector p_list = CZSystem.getCtTbRcp(roName1, g_no);
				if (p_list != null)
				{
					for(int i = 0 ; p_list.size() > i ; i++){
						dtTile = (CZSystemCtTitle)p_list.elementAt(i);
						cmbRcp1.addItem( dtTile.r_no + " : " + dtTile.title.trim());
						
						// @20131017 �O��I���������V�sNo���ύX�F�Ԑ�ɓ������V�sNo�����邩�H
						int icnt = cmbRcp1.getItemCount();
						CZSystem.log("CZTblHikaku2 setRcpDt","�R���{�{�b�N�X�� ���V�s���F" + icnt);
						CZSystem.log("CZTblHikaku2 setRcpDt","�R���{�{�b�N�X�� ���V�sNo�F" + dtTile.r_no);
						CZSystem.log("CZTblHikaku2 setRcpDt","�R���{�{�b�N�X�� �O��I�����V�sNo�F" + bufRcp1No);
						if (dtTile.r_no == bufRcp1No)  // �������V�sNo������ꍇ��index�擾
						{
							index1 = icnt;
							CZSystem.log("CZTblHikaku2 setRcpDt","�R���{�{�b�N�X Index���F" + index1);
						}
						// @20131017
					
					}
					//@20131017 �O��I�����V�s�ԍ����w��
					CZSystem.log("CZTblHikaku2 setRcpDt","�O��I�����V�s�ԍ�index1���w�� : " + index1);
					if (index1 != -1)
					{
						cmbRcp1.setSelectedIndex(index1-1);
					}
					// @20131017
				}

				Vector p_list2 = CZSystem.getCtTbRcp(roName2, g_no);
				if (p_list2 != null)
				{
					for(int i = 0 ; p_list2.size() > i ; i++){
						dtTile = (CZSystemCtTitle)p_list2.elementAt(i);
						cmbRcp2.addItem( dtTile.r_no + " : " + dtTile.title.trim());
						
						// @20131017 �O��I���������V�sNo���ύX�F�Ԑ�ɓ������V�sNo�����邩�H
						int icnt = cmbRcp2.getItemCount();
						CZSystem.log("CZTblHikaku2 setRcpDt","�R���{�{�b�N�X�� ���V�s���F" + icnt);
						CZSystem.log("CZTblHikaku2 setRcpDt","�R���{�{�b�N�X�� ���V�sNo�F" + dtTile.r_no);
						CZSystem.log("CZTblHikaku2 setRcpDt","�R���{�{�b�N�X�� �O��I�����V�sNo�F" + bufRcp2No);
						if (dtTile.r_no == bufRcp2No)  // �������V�sNo������ꍇ��index�擾
						{
							index2 = icnt;
							CZSystem.log("CZTblHikaku2 setRcpDt","�R���{�{�b�N�X Index2���F" + index2);
						}
						// @20131017
					
					}
					//@20131017 �O��I�����V�s�ԍ����w��
					CZSystem.log("CZTblHikaku2 setRcpDt","�O��I�����V�s�ԍ�index2���w�� : " + index2);
					if (index2 != -1)
					{
						cmbRcp2.setSelectedIndex(index2-1);
					}
					// @20131017
				}
			}
		}
	}
}
