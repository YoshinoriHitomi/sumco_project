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
 *   ���ƒ萔��r�pWindow 
 * @author  (KPK Co.,Ltd.)
 * @version 1.0 (2003/04/01)
 *
 */

public class CZTblHikaku extends JDialog {

    private RoNo    ro_from = null;
    private RoNo    ro_to = null;
    private int    iDispType;

    private JButton     hikaku_btn   = null;
    private JButton     cancel_button   = null;

    //
    //
    //
    CZTblHikaku(int iType){
        super();

CZSystem.log("CZTblHikaku ","iType=" + iType);
		if ((iType >= 1) && (iType <= 2)) {
			iDispType = iType;
		}
		else
		{
CZSystem.log("CZTblHikaku ","�`�o�G���[�\���^�C�v���ݒ肳��Ă��܂���");
			iDispType = 1;
		}

CZSystem.log("CZTblHikaku ","iDispType=" + iDispType);
        if (iDispType == 1)
			setTitle("���ƒ萔��r");
		else if (iDispType == 2)
    	    setTitle("����萔��r");
        
	
		setSize(280,200);
        setResizable(false);
        setModal(true);

        getContentPane().setLayout(null);
        // ����n�Q�Ƌ@�\    @20131021
        if(CZSystemDefine.REFERENCE_RUN == CZSystem.getRunLevel()){
            getContentPane().setBackground(CZSystemDefine.DEFAULT_REFERENCE_BACKGROUND_COL);
        }else{
            getContentPane().setBackground(CZSystemDefine.DEFAULT_BACKGROUND_COL);
        }

        JLabel  lab = new JLabel("��r��",JLabel.CENTER);
        lab.setBounds(20, 20, 100, 30);
        lab.setLocale(new Locale("ja","JP"));
        lab.setFont(new java.awt.Font("dialog", 0, 18));
        lab.setBorder(new Flush3DBorder());
        lab.setForeground(java.awt.Color.black);
        getContentPane().add(lab);

        ro_from = new RoNo();
        ro_from.setBounds(120, 20, 100, 30);
        getContentPane().add(ro_from);

        lab = new JLabel("��r��",JLabel.CENTER);
        lab.setBounds(20, 60, 100, 30);
        lab.setLocale(new Locale("ja","JP"));
        lab.setFont(new java.awt.Font("dialog", 0, 18));
        lab.setBorder(new Flush3DBorder());
        lab.setForeground(java.awt.Color.black);
        getContentPane().add(lab);

        ro_to = new RoNo();
        ro_to.setBounds(120, 60, 100, 30);
        getContentPane().add(ro_to);

        hikaku_btn = new JButton("��@�r");
        hikaku_btn.setBounds(20, 120, 100, 24);
        hikaku_btn.setLocale(new Locale("ja","JP"));
        hikaku_btn.setFont(new java.awt.Font("dialog", 0, 18));
        hikaku_btn.setBorder(new Flush3DBorder());
        hikaku_btn.setForeground(java.awt.Color.black);
        hikaku_btn.addActionListener(new hikaku_btn_click());
        getContentPane().add(hikaku_btn);

        cancel_button = new JButton("�I  ��");
        cancel_button.setBounds(140, 120, 100, 24);
        cancel_button.setLocale(new Locale("ja","JP"));
        cancel_button.setFont(new java.awt.Font("dialog", 0, 18));
        cancel_button.setBorder(new Flush3DBorder());
        cancel_button.setForeground(java.awt.Color.black);
        cancel_button.addActionListener(new CancelButton());
        getContentPane().add(cancel_button);
    }

    /*
    *
    *
    *
    */
    class hikaku_btn_click implements ActionListener {
        public void actionPerformed(ActionEvent ev){
			
	        Vector dataName = null;
	        Vector data = null;
	        Vector data2 = null;
			CZSystemOpTbAll dtName = null;
			CZSystemOpTb d = null;
			CZSystemOpTb d2 = null;
			int iRec1;
			int iRec2;
			int iNameRec;
			int iMax1;
			int iMax2;
			int iMidasi1 = 0;
			int iMidasi2 = 0;
			int iChkDtRtc;
			int iRt;

			String sHikakuHed = ",#,����,Min,Max,��,�P��,�l,�l,";

	        String sLine   = new String("");
	        String sDtOut  = new String("");

			//��r�������{

			String from_ro = ro_from.GetSelectedRoban();
			String to_ro = ro_to.GetSelectedRoban();
			
			CZSystem.log("hikaku_btn_click","from_ro=[" + from_ro + "] to_ro[" + to_ro + "]");
			
			dataName = CZSystem.opTblAllNameRead();
			data = CZSystem.getSogyoAllTb(ro_from.GetSelectedRobanDBname());
			data2 = CZSystem.getSogyoAllTb(ro_to.GetSelectedRobanDBname());

			String sNowDate = CZSystem.getDateTime("yyyy/MM/dd HH:mm:ss");
	        File file = new File(CZSystem.SOGYO_OUTPUT_PATH, "���Ɣ�r" + from_ro + "_"  + to_ro + "_"  + CZSystem.getDateTime("yyMMddHHmm") + ".csv");
	        PrintWriter pr     = null;
			FileOutputStream s = null;

	        try{
				s = new FileOutputStream(file);
				pr = new PrintWriter(s);

				sLine = "���ƒ萔�e�[�u����r,,,,,,,,," + sNowDate;
				pr.println(sLine);
				sLine = "";
				pr.println(sLine);
				sLine = "����������������r������������������";
				pr.println(sLine);
				sLine = "";
				pr.println(sLine);
				sLine = "�FNo1," + from_ro;
				pr.println(sLine);
				sLine = "�FNo2," + to_ro;
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
							dtName = (CZSystemOpTbAll)dataName.elementAt(iNameRec);
						}

						if (iRec1 < iMax1)
						{
							d = (CZSystemOpTb)data.elementAt(iRec1);
						}

						if (iRec2 < iMax2)
						{
							d2 = (CZSystemOpTb)data2.elementAt(iRec2);
						}

						/* ******* �f�[�^��r ******** */
						if ((iRec1 >= iMax1) && (iRec2 >= iMax2))
					    	iChkDtRtc = -99;
						else if (iRec1 >= iMax1)
					    	iChkDtRtc = 12;
						else if (iRec2 >= iMax2)
					    	iChkDtRtc = 11;
						else
					    	iChkDtRtc = chkDt(d, d2);

		CZSystem.log("hikaku_btn_click","iChkDtRtc = " + iChkDtRtc);

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
		CZSystem.log("hikaku_btn_click","iRt = " + iRt);

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
									
//									CZSystem.log("hikaku_btn_click","���o�� ptn1 [" + dtName.k_no1 + "][" + dtName.k_no2 + "]");
									iMidasi1 = dtName.k_no1;
									iMidasi2 = dtName.k_no2;
								}
								sDtOut = "," + dtName.k_no + "," + dtName.k_name.trim() + "," + dtName.n_min + "," + dtName.n_max + "," + dtName.keta + "," + dtName.t_name.trim() + ",";
//								CZSystem.log("hikaku_btn_click","���o�� [" + dtName.k_name + "]");
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
//								CZSystem.log("hikaku_btn_click","���o�� ptn1 [" + dtName.k_no1 + "][" + dtName.k_no2 + "]");
								iMidasi1 = dtName.k_no1;
								iMidasi2 = dtName.k_no2;
							}
							sDtOut = "," + dtName.k_no + "," + dtName.k_name.trim() + "," + dtName.n_min + "," + dtName.n_max + "," + dtName.keta + "," + dtName.t_name.trim() + ",-,-,";
							pr.println(sDtOut);
//							CZSystem.log("hikaku_btn_click","���o�� [" + dtName.k_name + "]");
//							CZSystem.log("hikaku_btn_click","dt [-][-]");
							iNameRec++;
						}
						else if (iRt == 2)
						{	/* �Y�����o���Ȃ� */
//							CZSystem.log("hikaku_btn_click","���o���Ȃ�");
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
//							CZSystem.log("hikaku_btn_click","����G���[");
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
//								CZSystem.log("hikaku_btn_click","ptn1 [" + d.k_no1 + "][" + d.k_no2 + "][" + d.k_no + "][" + d.k_val + "][" + d2.k_val + "]");
								iRec1++;	/* �f�[�^�P���R�[�hUP */
								iRec2++;	/* �f�[�^�Q���R�[�hUP */
							}
							else if (iChkDtRtc == 11)
							{	/* �f�[�^�Ⴄ�i�f�[�^�P�����Ȃ��j */
								sDtOut += d.k_val + ",-";
								pr.println(sDtOut);
//								CZSystem.log("hikaku_btn_click","ptn1 [" + d.k_no1 + "][" + d.k_no2 + "][" + d.k_no + "][" + d.k_val + "][ - ]");
								iRec1++;	/* �f�[�^�P���R�[�hUP */
							}
							else if (iChkDtRtc == 12)
							{	/* �f�[�^�Ⴄ�i�f�[�^�Q�����Ȃ��j */
								sDtOut += "-," + d2.k_val;
								pr.println(sDtOut);
//								CZSystem.log("hikaku_btn_click","ptn1 [" + d2.k_no1 + "][" + d2.k_no2 + "][" + d2.k_no + "][ - ][" + d2.k_val + "]");
								iRec2++;	/* �f�[�^�Q���R�[�hUP */
						    }
						}
					}	/* for end */
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
		                JOptionPane.showMessageDialog(null,"��r��̏�񂪂���܂���","��r����",JOptionPane.ERROR_MESSAGE);
					}
					else
					{
						CZSystem.log("hikaku_btn_click","��r�P�A�Q�f�[�^�Ȃ�");
						sLine = "";
						pr.println(sLine);
						sLine = "��r�P�A�Q�f�[�^�Ȃ�";
						pr.println(sLine);
		                JOptionPane.showMessageDialog(null,"��r��񂪂���܂���","��r����",JOptionPane.ERROR_MESSAGE);
					}
				}
	        }
	        catch(IOException e){
	            if(null != pr) pr.close();
	        }

	        if(null != pr) pr.close();
		}
    }

    private int chkDt(CZSystemOpTb dt,CZSystemOpTb dt2){
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

CZSystem.log("chkDt","rt=" + iRt + " [" + dt.k_no1 + "][" + dt.k_no2 + "][" + dt.k_no + "][" + dt.k_val + "] - [" + dt2.k_no1 + "][" + dt2.k_no2 + "][" + dt2.k_no + "][" + dt2.k_val + "]");
		return(iRt);
	}

    private int chkDtName(CZSystemOpTbAll dtName, CZSystemOpTb dt){
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

CZSystem.log("chkDtName","iRt=" + iRt + " [" + dt.k_no1 + "][" + dt.k_no2 + "][" + dt.k_no + "] - [" + dtName.k_no1 + "][" + dtName.k_no2 + "][" + dtName.k_no + "]");
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


    public boolean setType(int iType){
		iDispType = iType;
        return true;
    }

    //
    //
    //
    public class RoNo extends JComboBox {
        RoNo(){
            super();
            try{
                setName("JComboBox1");
                setFont(new java.awt.Font("dialog", 0, 16));
                Vector ro = CZSystem.getRoNameList();
                if(null == ro){
                    CZSystem.exit(0,"Not Ro No");
                }
                for(int i = 0 ; ro.size() > i ; i++){
					String s = CZSystem.RoKetaChg((String)ro.elementAt(i));	// 20050725 �F�F�\�������ύX
					addItem(s);
                }
                setForeground(java.awt.Color.black);
                setBackground(java.awt.Color.lightGray);
				setFocusable(false);	/* 2007.08.22 */
            }
            catch (Throwable e) {
                CZSystem.handleException(e);
            }
        }
		
		public String GetSelectedRoban() {
			return (String)getSelectedItem();
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
    } // RoNo

}
