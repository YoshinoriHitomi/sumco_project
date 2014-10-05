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
 *   操業定数比較用Window 
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
CZSystem.log("CZTblHikaku ","ＡＰエラー表示タイプが設定されていません");
			iDispType = 1;
		}

CZSystem.log("CZTblHikaku ","iDispType=" + iDispType);
        if (iDispType == 1)
			setTitle("操業定数比較");
		else if (iDispType == 2)
    	    setTitle("制御定数比較");
        
	
		setSize(280,200);
        setResizable(false);
        setModal(true);

        getContentPane().setLayout(null);
        // 他基地参照機能    @20131021
        if(CZSystemDefine.REFERENCE_RUN == CZSystem.getRunLevel()){
            getContentPane().setBackground(CZSystemDefine.DEFAULT_REFERENCE_BACKGROUND_COL);
        }else{
            getContentPane().setBackground(CZSystemDefine.DEFAULT_BACKGROUND_COL);
        }

        JLabel  lab = new JLabel("比較元",JLabel.CENTER);
        lab.setBounds(20, 20, 100, 30);
        lab.setLocale(new Locale("ja","JP"));
        lab.setFont(new java.awt.Font("dialog", 0, 18));
        lab.setBorder(new Flush3DBorder());
        lab.setForeground(java.awt.Color.black);
        getContentPane().add(lab);

        ro_from = new RoNo();
        ro_from.setBounds(120, 20, 100, 30);
        getContentPane().add(ro_from);

        lab = new JLabel("比較先",JLabel.CENTER);
        lab.setBounds(20, 60, 100, 30);
        lab.setLocale(new Locale("ja","JP"));
        lab.setFont(new java.awt.Font("dialog", 0, 18));
        lab.setBorder(new Flush3DBorder());
        lab.setForeground(java.awt.Color.black);
        getContentPane().add(lab);

        ro_to = new RoNo();
        ro_to.setBounds(120, 60, 100, 30);
        getContentPane().add(ro_to);

        hikaku_btn = new JButton("比　較");
        hikaku_btn.setBounds(20, 120, 100, 24);
        hikaku_btn.setLocale(new Locale("ja","JP"));
        hikaku_btn.setFont(new java.awt.Font("dialog", 0, 18));
        hikaku_btn.setBorder(new Flush3DBorder());
        hikaku_btn.setForeground(java.awt.Color.black);
        hikaku_btn.addActionListener(new hikaku_btn_click());
        getContentPane().add(hikaku_btn);

        cancel_button = new JButton("終  了");
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

			String sHikakuHed = ",#,項目,Min,Max,桁,単位,値,値,";

	        String sLine   = new String("");
	        String sDtOut  = new String("");

			//比較処理実施

			String from_ro = ro_from.GetSelectedRoban();
			String to_ro = ro_to.GetSelectedRoban();
			
			CZSystem.log("hikaku_btn_click","from_ro=[" + from_ro + "] to_ro[" + to_ro + "]");
			
			dataName = CZSystem.opTblAllNameRead();
			data = CZSystem.getSogyoAllTb(ro_from.GetSelectedRobanDBname());
			data2 = CZSystem.getSogyoAllTb(ro_to.GetSelectedRobanDBname());

			String sNowDate = CZSystem.getDateTime("yyyy/MM/dd HH:mm:ss");
	        File file = new File(CZSystem.SOGYO_OUTPUT_PATH, "操業比較" + from_ro + "_"  + to_ro + "_"  + CZSystem.getDateTime("yyMMddHHmm") + ".csv");
	        PrintWriter pr     = null;
			FileOutputStream s = null;

	        try{
				s = new FileOutputStream(file);
				pr = new PrintWriter(s);

				sLine = "操業定数テーブル比較,,,,,,,,," + sNowDate;
				pr.println(sLine);
				sLine = "";
				pr.println(sLine);
				sLine = "★★★★★★★比較条件★★★★★★★";
				pr.println(sLine);
				sLine = "";
				pr.println(sLine);
				sLine = "炉No1," + from_ro;
				pr.println(sLine);
				sLine = "炉No2," + to_ro;
				pr.println(sLine);
				sLine = "";
				pr.println(sLine);
				sLine = "★★★★★★★比較結果★★★★★★★";
				pr.println(sLine);
			
		        if ((null != dataName) &&(null != data) && (null != data2))
				{
		            iMax1 = data.size();
		            iMax2 = data2.size();
					iRec1 = 0;
					iRec2 = 0;
					iNameRec = 0;
		            for(int i = 0; ((iNameRec < dataName.size()) || (iRec1 < iMax1) || (iRec2 < iMax2)); i++){ 
						/* データチェック */
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

						/* ******* データ比較 ******** */
						if ((iRec1 >= iMax1) && (iRec2 >= iMax2))
					    	iChkDtRtc = -99;
						else if (iRec1 >= iMax1)
					    	iChkDtRtc = 12;
						else if (iRec2 >= iMax2)
					    	iChkDtRtc = 11;
						else
					    	iChkDtRtc = chkDt(d, d2);

		CZSystem.log("hikaku_btn_click","iChkDtRtc = " + iChkDtRtc);

						sDtOut = ",";	/* データ部クリア */
						iRt = -1;
						/* ******* 見出しチェック ******** */
						if (iNameRec >= dataName.size())
					    {
							iRt = 2;	/* 見出しなし */
						}
						else if ((iRec1 >= iMax1) && (iRec2 >= iMax2))
					    {
							iRt = 1;	/* 未設定項目出力(データ抜けの項目出力) */
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
								/* データ違う（データ２しかない） */
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
						/**************** 見出し出力 ****************/
						/********************************************/
						if (iRt == 0)
						{
							/* 見出し出力チェック */
							if (iChkDtRtc != 0)
							{
								if ((iMidasi1 != dtName.k_no1) || (iMidasi2 != dtName.k_no2))
								{
									sLine = "";
									pr.println(sLine);
									sLine = "【" + dtName.k_name1.trim() + "】 ： 【" + dtName.k_name2.trim() + "】";
									pr.println(sLine);
									sLine = sHikakuHed;
									pr.println(sLine);
									
//									CZSystem.log("hikaku_btn_click","見出し ptn1 [" + dtName.k_no1 + "][" + dtName.k_no2 + "]");
									iMidasi1 = dtName.k_no1;
									iMidasi2 = dtName.k_no2;
								}
								sDtOut = "," + dtName.k_no + "," + dtName.k_name.trim() + "," + dtName.n_min + "," + dtName.n_max + "," + dtName.keta + "," + dtName.t_name.trim() + ",";
//								CZSystem.log("hikaku_btn_click","見出し [" + dtName.k_name + "]");
							}
							iNameRec++;	/* 項目名称レコードチェンジ */
						}
						else if (iRt == 1)
						{	/* 未設定項目出力(データ抜けの項目出力) */
							if ((iMidasi1 != dtName.k_no1) || (iMidasi2 != dtName.k_no2))
							{
								sLine = "";
								pr.println(sLine);
								sLine = "【" + dtName.k_name1.trim() + "】 ： 【" + dtName.k_name2.trim() + "】";
								pr.println(sLine);
								sLine = sHikakuHed;
								pr.println(sLine);
//								CZSystem.log("hikaku_btn_click","見出し ptn1 [" + dtName.k_no1 + "][" + dtName.k_no2 + "]");
								iMidasi1 = dtName.k_no1;
								iMidasi2 = dtName.k_no2;
							}
							sDtOut = "," + dtName.k_no + "," + dtName.k_name.trim() + "," + dtName.n_min + "," + dtName.n_max + "," + dtName.keta + "," + dtName.t_name.trim() + ",-,-,";
							pr.println(sDtOut);
//							CZSystem.log("hikaku_btn_click","見出し [" + dtName.k_name + "]");
//							CZSystem.log("hikaku_btn_click","dt [-][-]");
							iNameRec++;
						}
						else if (iRt == 2)
						{	/* 該当見出しなし */
//							CZSystem.log("hikaku_btn_click","見出しなし");
							if (iChkDtRtc == 12)
							{
								if ((iMidasi1 != d2.k_no1) || (iMidasi2 != d2.k_no2))
								{
									sLine = "";
									pr.println(sLine);
									sLine = "【" + d2.k_no1 + "】 ： 【" + d2.k_no2 + "】";
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
									sLine = "【" + d.k_no1 + "】 ： 【" + d.k_no2 + "】";
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
//							CZSystem.log("hikaku_btn_click","判定エラー");
							sDtOut = "判定エラー,-,-,-,-,-,-,";
							iNameRec++;
							pr.println(sDtOut);
						}

						/********************************************/
						/**************** データ出力 ****************/
						/********************************************/
						if ((iRt == 0) || (iRt == 2))
						{
							if (iChkDtRtc == 0)
							{	/* データ同じ */
								/* 例外処理 */
								/* 定義にない項目の場合は、データが同じでも */
								/* 値を表示 */
								if (iRt == 2) 
								{
									sDtOut += d.k_val + "," + d2.k_val;
									pr.println(sDtOut);
								}
								iRec1++;	/* データ１レコードUP */
								iRec2++;	/* データ２レコードUP */
							}
							else if (iChkDtRtc == 1)
							{	/* データ違う（値が違う） */
								sDtOut += d.k_val + "," + d2.k_val;
								pr.println(sDtOut);
//								CZSystem.log("hikaku_btn_click","ptn1 [" + d.k_no1 + "][" + d.k_no2 + "][" + d.k_no + "][" + d.k_val + "][" + d2.k_val + "]");
								iRec1++;	/* データ１レコードUP */
								iRec2++;	/* データ２レコードUP */
							}
							else if (iChkDtRtc == 11)
							{	/* データ違う（データ１しかない） */
								sDtOut += d.k_val + ",-";
								pr.println(sDtOut);
//								CZSystem.log("hikaku_btn_click","ptn1 [" + d.k_no1 + "][" + d.k_no2 + "][" + d.k_no + "][" + d.k_val + "][ - ]");
								iRec1++;	/* データ１レコードUP */
							}
							else if (iChkDtRtc == 12)
							{	/* データ違う（データ２しかない） */
								sDtOut += "-," + d2.k_val;
								pr.println(sDtOut);
//								CZSystem.log("hikaku_btn_click","ptn1 [" + d2.k_no1 + "][" + d2.k_no2 + "][" + d2.k_no + "][ - ][" + d2.k_val + "]");
								iRec2++;	/* データ２レコードUP */
						    }
						}
					}	/* for end */
	                JOptionPane.showMessageDialog(null,"比較が完了しました。","比較処理",JOptionPane.INFORMATION_MESSAGE);
		        }
				else
				{
			        if ((null == data) && (null != data2))
					{
						CZSystem.log("hikaku_btn_click","比較１データなし");
						sLine = "";
						pr.println(sLine);
						sLine = "比較１データなし";
						pr.println(sLine);
		                JOptionPane.showMessageDialog(null,"比較元の情報がありません","比較処理",JOptionPane.ERROR_MESSAGE);
					}
					else
			        if ((null != data) && (null == data2))
					{
						CZSystem.log("hikaku_btn_click","比較２データなし");
						sLine = "";
						pr.println(sLine);
						sLine = "比較２データなし";
						pr.println(sLine);
		                JOptionPane.showMessageDialog(null,"比較先の情報がありません","比較処理",JOptionPane.ERROR_MESSAGE);
					}
					else
					{
						CZSystem.log("hikaku_btn_click","比較１、２データなし");
						sLine = "";
						pr.println(sLine);
						sLine = "比較１、２データなし";
						pr.println(sLine);
		                JOptionPane.showMessageDialog(null,"比較情報がありません","比較処理",JOptionPane.ERROR_MESSAGE);
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
		int  iRt = -1;	/* 項目名チェック結果　0:同じ値　*/
						/*                     1:違う値　*/
						/*					　 11:データ１のみデータあり */
						/*					　 12:データ２のみデータあり */
						/*					　 -1:判定異常 */

		/* 項目名称検索 */
		if (dt.k_no1 == dt2.k_no1)
		{
			if (dt.k_no2 == dt2.k_no2)
			{
				if(dt.k_no == dt2.k_no)
				{
					/* 同一項目あり */
					if (dt.k_val == dt2.k_val)
					{
						/* データ同じ */
						iRt = 0;
					}
					else
					{
						/* データ違う */
						iRt = 1;
					}
				}
				else if (dt.k_no < dt2.k_no)
				{
					/* データ１データあり */
					iRt = 11;
				}
				else
				{
					/* データ２データあり */
					iRt = 12;
				}
			}
			else if (dt.k_no2 < dt2.k_no2)
			{
				/* データ１データあり */
				iRt = 11;
			}
			else
			{
				/* データ２データあり */
				iRt = 12;
			}
		}
		else if (dt.k_no1 < dt2.k_no1)
		{
			/* データ１データあり */
			iRt = 11;
		}
		else
		{
			/* データ２データあり */
			iRt = 12;
		}

CZSystem.log("chkDt","rt=" + iRt + " [" + dt.k_no1 + "][" + dt.k_no2 + "][" + dt.k_no + "][" + dt.k_val + "] - [" + dt2.k_no1 + "][" + dt2.k_no2 + "][" + dt2.k_no + "][" + dt2.k_val + "]");
		return(iRt);
	}

    private int chkDtName(CZSystemOpTbAll dtName, CZSystemOpTb dt){
		int  iRt = -1;	/* 項目名チェック結果　0:同一項目あり　*/
						/*                     1:炉データ未設定情報あり（定義のみあり）　*/
						/*					　 2:該当項目名なし */
						/*					　 -1:判定異常 */

		/* 項目名称検索 */
		if (dtName.k_no1 == dt.k_no1)
		{
			if (dtName.k_no2 == dt.k_no2)
			{
				if(dtName.k_no == dt.k_no)
				{
					/* 同一項目あり */
					/* 名称出力 */
					/* 項目インクリメント */
					iRt = 0;
				}
				else if (dtName.k_no < dt.k_no)
				{
					/* 名称出力（データなし） */
					/* 項目インクリメント */
					iRt = 1;
				}
				else
				{
					/* 名称なし */
					iRt = 2;
				}
			}
			else if (dtName.k_no2 < dt.k_no2)
			{
				/* 名称出力（データなし） */
				/* 項目インクリメント */
				iRt = 1;
			}
			else
			{
				/* 名称なし */
				iRt = 2;
			}
		}
		else if (dtName.k_no1 < dt.k_no1)
		{
			/* 名称出力（データなし） */
			/* 項目インクリメント */
			iRt = 1;
		}
		else
		{
			/* 名称なし */
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
					String s = CZSystem.RoKetaChg((String)ro.elementAt(i));	// 20050725 炉：表示桁数変更
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
