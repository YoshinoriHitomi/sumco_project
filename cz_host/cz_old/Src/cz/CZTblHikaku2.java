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
 *   制御テーブル比較用Window 
 * @author  (KPK Co.,Ltd.)
 * @version 1.0 (2003/04/01)
 * @Update 2013/10/17 レシピ番号保持機能 (@20131017)
 */

public class CZTblHikaku2 extends JDialog {

    private RoNo    ro_from = null;
    private RoNo    ro_to = null;

    private GrupNo  cmdGrup = null;

    private JComboBox    cmbRcp1 = null;
    private JComboBox    cmbRcp2 = null;

    private JButton     hikaku_btn   = null;
    private JButton     cancel_button   = null;

	String sHikakuHed = ",#,項目,Min,Max,桁,単位,値,値,";
    String sLine   = new String("");
    String sDtOut  = new String("");

    // 比較元
    String sRcp1Info = new String("");  //@20131017
    int bufRcp1No = 0;  //@20131017

    // 比較先
    String sRcp2Info = new String("");  //@20131017
    int bufRcp2No = 0;  //@20131017

    //
    //
    //
    CZTblHikaku2(){
        super();

   	    setTitle("制御テーブル比較");
	
		setSize(460,320);
        setResizable(false);
        setModal(true);

        getContentPane().setLayout(null);
        // 他基地参照機能    @20131021
        if(CZSystemDefine.REFERENCE_RUN == CZSystem.getRunLevel()){
            getContentPane().setBackground(CZSystemDefine.DEFAULT_REFERENCE_BACKGROUND_COL);
        }else{
            getContentPane().setBackground(CZSystemDefine.DEFAULT_BACKGROUND_COL);
        }


        JLabel  lab = new JLabel("グループ",JLabel.CENTER);
        lab.setBounds(20, 20, 100, 30);
        lab.setLocale(new Locale("ja","JP"));
        lab.setFont(new java.awt.Font("dialog", 0, 18));
        lab.setBorder(new Flush3DBorder());
        lab.setForeground(java.awt.Color.black);
        getContentPane().add(lab);

        cmdGrup = new GrupNo();
        cmdGrup.setBounds(120, 20, 100, 30);
        getContentPane().add(cmdGrup);


        lab = new JLabel("比較元",JLabel.CENTER);
        lab.setBounds(20, 70, 100, 30);
        lab.setLocale(new Locale("ja","JP"));
        lab.setFont(new java.awt.Font("dialog", 0, 18));
        lab.setBorder(new Flush3DBorder());
        lab.setForeground(java.awt.Color.black);
        getContentPane().add(lab);

        ro_from = new RoNo(1);
        ro_from.setBounds(120, 70, 100, 30);
        getContentPane().add(ro_from);


        lab = new JLabel("レシピNo",JLabel.CENTER);
        lab.setBounds(20, 100, 100, 30);
        lab.setLocale(new Locale("ja","JP"));
        lab.setFont(new java.awt.Font("dialog", 0, 18));
        lab.setBorder(new Flush3DBorder());
        lab.setForeground(java.awt.Color.black);
        getContentPane().add(lab);

        cmbRcp1 = new JComboBox();
        cmbRcp1.setBounds(120, 100, 300, 30);
		cmbRcp1.setFocusable(false);	/* 2007.08.22 */

		//ＴＯ炉選択中レシピ格納
		ro_from.setRcpDt();
        getContentPane().add(cmbRcp1);

        lab = new JLabel("比較先",JLabel.CENTER);
        lab.setBounds(20, 140, 100, 30);
        lab.setLocale(new Locale("ja","JP"));
        lab.setFont(new java.awt.Font("dialog", 0, 18));
        lab.setBorder(new Flush3DBorder());
        lab.setForeground(java.awt.Color.black);
        getContentPane().add(lab);

        ro_to = new RoNo(2);
        ro_to.setBounds(120, 140, 100, 30);
        getContentPane().add(ro_to);


        lab = new JLabel("レシピNo",JLabel.CENTER);
        lab.setBounds(20, 170, 100, 30);
        lab.setLocale(new Locale("ja","JP"));
        lab.setFont(new java.awt.Font("dialog", 0, 18));
        lab.setBorder(new Flush3DBorder());
        lab.setForeground(java.awt.Color.black);
        getContentPane().add(lab);

        cmbRcp2 = new JComboBox();
        cmbRcp2.setBounds(120, 170, 300, 30);
		cmbRcp2.setFocusable(false);	/* 2007.08.22 */
		//ＴＯ炉選択中レシピ格納
		ro_to.setRcpDt();
        getContentPane().add(cmbRcp2);


        hikaku_btn = new JButton("比　較");
        hikaku_btn.setBounds(20, 240, 100, 24);
        hikaku_btn.setLocale(new Locale("ja","JP"));
        hikaku_btn.setFont(new java.awt.Font("dialog", 0, 18));
        hikaku_btn.setBorder(new Flush3DBorder());
        hikaku_btn.setForeground(java.awt.Color.black);
        hikaku_btn.addActionListener(new hikaku_btn_click());
        getContentPane().add(hikaku_btn);

        cancel_button = new JButton("終  了");
        cancel_button.setBounds(140, 240, 100, 24);
        cancel_button.setLocale(new Locale("ja","JP"));
        cancel_button.setFont(new java.awt.Font("dialog", 0, 18));
        cancel_button.setBorder(new Flush3DBorder());
        cancel_button.setForeground(java.awt.Color.black);
        cancel_button.addActionListener(new CancelButton());
        getContentPane().add(cancel_button);
    }


    //
    // メッセージの表示
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
			
			int tblGno = cmdGrup.getGrupNo();				//選択グループＮｏ取得
			String from_ro = ro_from.GetSelectedRoban();	//選択炉名（比較元）　取得
			String to_ro = ro_to.GetSelectedRoban();		//選択炉名（比較先）　取得
			String sBuf = null;								//変換用仮バッファー
			int rcp_from = 0;								//選択レシピ（比較元）　
			int rcp_to = 0;									//選択レシピ（比較先）　


			//選択レシピＮｏ（比較元）取得
			sBuf = (String)cmbRcp1.getSelectedItem();
			if (sBuf != null)
			{
				if (sBuf.indexOf(" ") != -1)
					rcp_from = Integer.valueOf(sBuf.substring(0,sBuf.indexOf(" "))).intValue();
			}
			else
			{
                JOptionPane.showMessageDialog(null,"比較元のレシピNoなし","比較処理",JOptionPane.ERROR_MESSAGE);
//errorMsg("比較処理", "比較元のレシピNoなし");
				CZSystem.log("hikaku_btn_click","レシピなし（１）");
				return;
			}

			//選択レシピＮｏ（比較先）取得
			sBuf = (String)cmbRcp2.getSelectedItem();
			if (sBuf != null)
			{
				if (sBuf.indexOf(" ") != -1)
					rcp_to = Integer.valueOf(sBuf.substring(0,sBuf.indexOf(" "))).intValue();
			}
			else
			{
//                errorMsg("比較処理","比較先のレシピNoなし");
                JOptionPane.showMessageDialog(null,"比較先のレシピNoなし","比較処理",JOptionPane.ERROR_MESSAGE);
				CZSystem.log("hikaku_btn_click","レシピなし（２）");
				return;
			}
			

			CZSystem.log("hikaku_btn_click","from_ro=[" + from_ro + "] to_ro[" + to_ro + "]");
			
			//選択グループ判定
			if (tblGno == 6)
			{
				//比較処理実施
				subT6Chk(ro_from.GetSelectedRobanDBname(), from_ro, rcp_from, 
						 ro_to.GetSelectedRobanDBname(),   to_ro,   rcp_to);
			}
			else
			{
				//比較処理実施
				subT1_5Chk(tblGno, ro_from.GetSelectedRobanDBname(), from_ro, rcp_from, 
						 ro_to.GetSelectedRobanDBname(),   to_ro,   rcp_to);
			}
		}
   }

    private void subT1_5Chk(int tblGno, String fromDB_ro, String from_ro, int rcp_from, String toDB_ro, String to_ro, int rcp_to){
//		int		tblGno;		選択されたグループNo
//		String fromDB_ro;	選択炉名（比較元のＤＢ名称）
//		String from_ro;		選択炉名（比較元の表示名称）
//		int rcp_from;		選択レシピ（比較元）
//		String toDB_ro;		選択炉名（比較先のＤＢ名称）
//		String to_ro;		選択炉名（比較先の表示名称）
//		int rcp_to;			選択レシピ（比較先）
	
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


		String sBuf = null;								//変換用仮バッファー
		
		//名称取得
		dataName = CZSystem.ctTblAllNameRead(tblGno);

		//比較元のレシピ情報取得
		data = CZSystem.getCtAllTb(fromDB_ro, tblGno, rcp_from);

		//比較先のＴ６レシピ情報取得
		data2 = CZSystem.getCtAllTb(toDB_ro, tblGno, rcp_to);

		//現在時刻取得
		String sNowDate = CZSystem.getDateTime("yyyy/MM/dd HH:mm:ss");

		//ファイル生成
        File file = new File(CZSystem.SOGYO_OUTPUT_PATH, "制御比較" + from_ro + "_"  + to_ro + "_"  + CZSystem.getDateTime("yyMMddHHmm") + ".csv");
        PrintWriter pr     = null;
		FileOutputStream s = null;

        try{
			s = new FileOutputStream(file);
			pr = new PrintWriter(s);

			sLine = "制御テーブル比較（ T" + tblGno + " ）,,,,,,,,," + sNowDate;
			pr.println(sLine);
			sLine = "";
			pr.println(sLine);
			sLine = "★★★★★★★比較条件★★★★★★★";
			pr.println(sLine);
			sLine = "";
			pr.println(sLine);
			sLine = "炉No1," + from_ro + ",";
			pr.println(sLine);
			sLine = "レシピNo," + rcp_from + ",";
			pr.println(sLine);
			sLine = "";
			pr.println(sLine);

			sLine = "炉No2," + to_ro + ",";
			pr.println(sLine);
			sLine = "レシピNo," + rcp_to + ",";
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

					/* ******* データ比較 ******** */
					if ((iRec1 >= iMax1) && (iRec2 >= iMax2))
				    {
						/* 見出しのみ出力 */
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

										//全レコードありor見出しなし

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


						/* 項目名称検索 */
						if (dt.t_no == dt2.t_no)
						{
							tno = dt.t_no;
							iChkDtRtc = 0;
			CZSystem.log("データチェックST","[" + iRecChk1 + "] [" + iMax1 + "]" + "[" + iRecChk2 + "] [" + iMax2 + "]");
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

								//値のチェック 2008.01.28
								if (((tno != dt.t_no) && (tno != dt2.t_no)) ||
								    ((iRecChk1 >= iMax1) && (tno != dt2.t_no)) ||
								    ((tno != dt.t_no) && (iRecChk2 >= iMax2)))
								{	//両方違う（チェック終了）
									break;
								}

								if (((tno != dt.t_no) && (tno == dt2.t_no)) ||
								    ((tno == dt.t_no) && (tno != dt2.t_no)))
								{
									/* データ違う */
									iChkDtRtc = 1;
									break;
								}

								if ((tno == dt.t_no) && (tno == dt2.t_no))
								{
									/* 同一項目あり */
									if ((dt.k_no == dt2.k_no) && 
									    (dt.l_val == dt2.l_val) && 
									    (dt.r_val == dt2.r_val))
									{
										/* データ同じ */
									}
									else
									{
										/* データ違う */
										iChkDtRtc = 1;
										break;
									}
								}
								iRecChk1++;
								iRecChk2++;
			CZSystem.log("データチェック中","[" + iRecChk1 + "] [" + iMax1 + "]" + "[" + iRecChk2 + "] [" + iMax2 + "]");
							}
			CZSystem.log("データチェック終了","[" + iRecChk1 + "] [" + iMax1 + "]" + "[" + iRecChk2 + "] [" + iMax2 + "]");

						}
						else if (dt.t_no < dt2.t_no)
						{
							/* データ１データあり */
							iChkDtRtc = 11;
						}
						else
						{
							/* データ２データあり */
							iChkDtRtc = 12;
						}
					}

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

					/********************************************/
					/**************** 見出し出力 ****************/
					/********************************************/
					if (iRt == 0)
					{
						/* 見出し出力チェック */
						if (iChkDtRtc != 0)
						{
							if (iMidasi1 != dtName.t_no)
							{
								sLine = "";
								pr.println(sLine);
								sLine = "【" + dtName.t_no + " ： " + dtName.t_name.trim() + "】";
								pr.println(sLine);

								iMidasi1 = dtName.t_no;
							}
							sDtOut = ",Ｌ軸,Ｒ軸,,Ｌ軸,Ｒ軸,";
							pr.println(sDtOut);
						}
						iNameRec++;	/* 項目名称レコードチェンジ */
					}
					else if (iRt == 1)
					{	/* 未設定項目出力(データ抜けの項目出力) */
						if (iMidasi1 != dtName.t_no)
						{
							sLine = "";
							pr.println(sLine);
							sLine = "【" + dtName.t_no + " ： " + dtName.t_name.trim() + "】";
							pr.println(sLine);
							
							iMidasi1 = dtName.t_no;
						}
						sDtOut = ",Ｌ軸,Ｒ軸,,Ｌ軸,Ｒ軸,";
						pr.println(sDtOut);
						iNameRec++;	/* 項目名称レコードチェンジ */
					}
					else if (iRt == 2)
					{	/* 該当見出しなし */
						if (iChkDtRtc == 12)
						{
							if (iMidasi1 != d2.t_no)
							{
								sLine = "";
								pr.println(sLine);
								sLine = "【" + d2.t_no + "】";
								pr.println(sLine);

								iMidasi1 = d2.t_no;
							}
							sDtOut = ",Ｌ軸,Ｒ軸,,Ｌ軸,Ｒ軸,";
							pr.println(sDtOut);
						}
						else
						{
							if (iMidasi1 != d.t_no)
							{
								sLine = "";
								pr.println(sLine);
								sLine = "【" + d.t_no + "】";
								pr.println(sLine);

								iMidasi1 = d.t_no;
							}
							sDtOut = ",Ｌ軸,Ｒ軸,,Ｌ軸,Ｒ軸,";
							pr.println(sDtOut);
						}
					}
					else
					{
						sDtOut = "判定エラー,-,-,-,-,-,-,";
						iNameRec++;
						pr.println(sDtOut);
					}

					/********************************************/
					/**************** データ出力 ****************/
					/********************************************/
					if ((iRt == 0) || (iRt == 2))
					{
						tno = d.t_no;
						if (iChkDtRtc == 12)
							tno = d2.t_no;

			CZSystem.log("データ出力開始","[" + iRec1 + "] [" + iMax1 + "]" + "[" + iRec2 + "] [" + iMax2 + "]");
						for (int iLp=0; ((iRec1 < iMax1) || (iRec2 < iMax2));iLp++)
						{
			CZSystem.log("データ出力--start","[" + iChkDtRtc + "] [" + iRt + "]");
							if (iLp >= 32760)
							{
			CZSystem.log("データ出力--強制終了","Loop　OVER");
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
								break;	//ターゲット変更
							}

							if (iChkDtRtc == 0)
							{	/* データ同じ */
								/* 例外処理 */
								/* 定義にない項目の場合は、データが同じでも */
								/* 値を表示 */
								if (iRt == 2) 
								{
									sDtOut = ",";
									if ((iRec1 < iMax1) && (tno == d.t_no))
									{
										sDtOut += d.l_val + "," + d.r_val + ",";
										iRec1++;	/* データ１レコードUP */
									}
									else
										sDtOut += "-,-,";

									if ((iRec2 < iMax2) && (tno == d2.t_no))
									{
										sDtOut += "," + d2.l_val + "," + d2.r_val + ",";
										iRec2++;	/* データ２レコードUP */
									}
									else
										sDtOut += ",-,-,";

									pr.println(sDtOut);
								}
								else
								{
									iRec1++;	/* データ１レコードUP */
									iRec2++;	/* データ２レコードUP */
								}


							}
							else if (iChkDtRtc == 1)
							{	/* データ違う（値が違う） */
								//値のチェック
								sDtOut = ",";
								if ((iRec1 < iMax1) && (tno == d.t_no))
								{
									sDtOut += d.l_val + "," + d.r_val + ",";
									iRec1++;	/* データ１レコードUP */
								}
								else
									sDtOut += "-,-,";

								if ((iRec2 < iMax2) && (tno == d2.t_no))
								{
									sDtOut += "," + d2.l_val + "," + d2.r_val + ",";
									iRec2++;	/* データ２レコードUP */
								}
								else
									sDtOut += ",-,-,";

								pr.println(sDtOut);
							}
							else if (iChkDtRtc == 11)
							{	/* データ違う（データ１しかない） */
								sDtOut = ",";
								if ((iRec1 < iMax1) && (tno == d.t_no))
								{
									sDtOut += d.l_val + "," + d.r_val + ",";
									iRec1++;	/* データ１レコードUP */
								}
								else
									sDtOut += "-,-,";

								sDtOut += ",-,-,";

								pr.println(sDtOut);
							}
							else if (iChkDtRtc == 12)
							{	/* データ違う（データ２しかない） */
								sDtOut = ",";
								sDtOut += "-,-,";

								if ((iRec2 < iMax2) && (tno == d2.t_no))
								{
									sDtOut += "," + d2.l_val + "," + d2.r_val + ",";
									iRec2++;	/* データ２レコードUP */
								}
								else
									sDtOut += ",-,-,";

								pr.println(sDtOut);
						    }
						}
			CZSystem.log("データ出力中","[" + iRec1 + "] [" + iMax1 + "]" + "[" + iRec2 + "] [" + iMax2 + "]");
					}
				}	/* for end */
			CZSystem.log("データ出力完了","[" + iRec1 + "] [" + iMax1 + "]" + "[" + iRec2 + "] [" + iMax2 + "]");
//                infoMsg("比較処理", "比較完了");
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
//	                errorMsg("比較処理","比較元の情報がありません");
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
//	                errorMsg("比較処理","比較先の情報がありません");
	                JOptionPane.showMessageDialog(null,"比較先の情報がありません","比較処理",JOptionPane.ERROR_MESSAGE);
				}
				else
				{
					CZSystem.log("hikaku_btn_click","比較１、２データなし");
					sLine = "";
					pr.println(sLine);
					sLine = "比較１、２データなし";
					pr.println(sLine);
//	                errorMsg("比較処理","比較情報がありません");
	                JOptionPane.showMessageDialog(null,"比較情報がありません","比較処理",JOptionPane.ERROR_MESSAGE);
				}
			}
        }
        catch(IOException e){
            if(null != pr) pr.close();
        }

        if(null != pr) pr.close();
	}

    private void subT6Chk(String fromDB_ro, String from_ro, int rcp_from, String toDB_ro, String to_ro, int rcp_to){
//		String fromDB_ro;	選択炉名（比較元のＤＢ名称）
//		String from_ro;		選択炉名（比較元の表示名称）
//		int rcp_from;		選択レシピ（比較元）
//		String toDB_ro;		選択炉名（比較先のＤＢ名称）
//		String to_ro;		選択炉名（比較先の表示名称）
//		int rcp_to;			選択レシピ（比較先）
	
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


		String sBuf = null;								//変換用仮バッファー
		
		//Ｔ６名称取得
		dataName = CZSystem.ctT6AllNameRead();

		//比較元のＴ６レシピ情報取得
		data = CZSystem.getCtT6AllTb(fromDB_ro,rcp_from);

		//比較先のＴ６レシピ情報取得
		data2 = CZSystem.getCtT6AllTb(toDB_ro,rcp_to);

		//現在時刻取得
		String sNowDate = CZSystem.getDateTime("yyyy/MM/dd HH:mm:ss");

		//ファイル生成
        File file = new File(CZSystem.SOGYO_OUTPUT_PATH, "制御比較" + from_ro + "_"  + to_ro + "_"  + CZSystem.getDateTime("yyMMddHHmm") + ".csv");
        PrintWriter pr     = null;
		FileOutputStream s = null;

        try{
			s = new FileOutputStream(file);
			pr = new PrintWriter(s);

			sLine = "制御テーブル比較（Ｔ６）,,,,,,,,," + sNowDate;
			pr.println(sLine);
			sLine = "";
			pr.println(sLine);
			sLine = "★★★★★★★比較条件★★★★★★★";
			pr.println(sLine);
			sLine = "";
			pr.println(sLine);
			sLine = "炉No1," + from_ro + ",";
			pr.println(sLine);
			sLine = "レシピNo," + rcp_from + ",";
			pr.println(sLine);
			sLine = "";
			pr.println(sLine);

			sLine = "炉No2," + to_ro + ",";
			pr.println(sLine);
			sLine = "レシピNo," + rcp_to + ",";
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

					/* ******* データ比較 ******** */
					if ((iRec1 >= iMax1) && (iRec2 >= iMax2))
				    	iChkDtRtc = -99;
					else if (iRec1 >= iMax1)
				    	iChkDtRtc = 12;
					else if (iRec2 >= iMax2)
				    	iChkDtRtc = 11;
					else
				    	iChkDtRtc = chkDtT6(d, d2);

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
					    	iRt = chkDtNameT6(dtName, d);
						}
						else if (iChkDtRtc == 12)
						{
							/* データ違う（データ２しかない） */
					    	iRt = chkDtNameT6(dtName, d2);
					    }
						else
						{
							CZSystem.log("hikaku_btn_click","chkDtRtc err sts [" + iChkDtRtc + "]");
							break;
						}
					}

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
								
								iMidasi1 = dtName.k_no1;
								iMidasi2 = dtName.k_no2;
							}
							sDtOut = "," + dtName.k_no + "," + dtName.k_name.trim() + "," + dtName.k_min + "," + dtName.k_max + "," + dtName.k_keta + "," + dtName.k_unit.trim() + ",";
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
							iMidasi1 = dtName.k_no1;
							iMidasi2 = dtName.k_no2;
						}
						sDtOut = "," + dtName.k_no + "," + dtName.k_name.trim() + "," + dtName.k_min + "," + dtName.k_max + "," + dtName.k_keta + "," + dtName.k_unit.trim() + ",-,-,";
						pr.println(sDtOut);
						iNameRec++;
					}
					else if (iRt == 2)
					{	/* 該当見出しなし */
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
							iRec1++;	/* データ１レコードUP */
							iRec2++;	/* データ２レコードUP */
						}
						else if (iChkDtRtc == 11)
						{	/* データ違う（データ１しかない） */
							sDtOut += d.k_val + ",-";
							pr.println(sDtOut);
							iRec1++;	/* データ１レコードUP */
						}
						else if (iChkDtRtc == 12)
						{	/* データ違う（データ２しかない） */
							sDtOut += "-," + d2.k_val;
							pr.println(sDtOut);
							iRec2++;	/* データ２レコードUP */
					    }
					}
				}	/* for end */
//                infoMsg("比較処理","比較完了");
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
//	                errorMsg("比較処理","比較元の情報がありません");
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
//	                errorMsg("比較処理","比較先の情報がありません");
	                JOptionPane.showMessageDialog(null,"比較先の情報がありません","比較処理",JOptionPane.ERROR_MESSAGE);
				}
				else
				{
					CZSystem.log("hikaku_btn_click","比較１、２データなし");
					sLine = "";
					pr.println(sLine);
					sLine = "比較１、２データなし";
					pr.println(sLine);
//	                errorMsg("比較処理","比較情報がありません");
	                JOptionPane.showMessageDialog(null,"比較情報がありません","比較処理",JOptionPane.ERROR_MESSAGE);
				}
			}
        }
        catch(IOException e){
            if(null != pr) pr.close();
        }

        if(null != pr) pr.close();
	}



	//Ｔ６データ比較
    private int chkDtT6(CZSystemCtT6Tb dt,CZSystemCtT6Tb dt2){
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

		return(iRt);
	}

	//Ｔ１～Ｔ５見出しチェック
    private int chkDtName(CZSystemCtName dtName, CZSystemCtTb dt){
		int  iRt = -1;	/* 項目名チェック結果　0:同一項目あり　*/
						/*                     1:炉データ未設定情報あり（定義のみあり）　*/
						/*					　 2:該当項目名なし */
						/*					　 -1:判定異常 */

		/* 項目名称検索 */
		if (dtName.t_no == dt.t_no)
		{
			/* 同一項目あり */
			/* 名称出力 */
			/* 項目インクリメント */
			iRt = 0;
		}
		else if (dtName.t_no < dt.t_no)
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

		return(iRt);
	}

	//Ｔ６見出しチェック
    private int chkDtNameT6(CZSystemCtT6AllName dtName, CZSystemCtT6Tb dt){
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
	// 炉番を格納する
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
					String s = CZSystem.RoKetaChg((String)ro.elementAt(i));	// 20050725 炉：表示桁数変更
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

		//現在選択中の炉番のレシピを格納
		public void setRcpDt() {
			String roName;
			int g_no;
			int index1 = -1;  // @20131017
			int index2 = -1;  // @20131017
			CZSystemCtTitle   dtTile;
			roName = GetSelectedRobanDBname();

			g_no = cmdGrup.getGrupNo();

			CZSystem.log("setRcpDt", "炉番 : " + roName + "g_no : " + g_no + "モード[" + iMode + "]");

			if (iMode == 1){
				//選択レシピＮｏ（比較元）取得 @20131017
				sRcp1Info = (String)cmbRcp1.getSelectedItem();
				if (sRcp1Info != null)
				{
					if (sRcp1Info.indexOf(" ") != -1)
						bufRcp1No = Integer.valueOf(sRcp1Info.substring(0,sRcp1Info.indexOf(" "))).intValue();
				}
	            CZSystem.log("CZTblHikaku2 ChgRcpNo","（比較元）レシピ番号変更");
	            CZSystem.log("CZTblHikaku2 ChgRcpNo","（比較元）レシピ番号取得："+ bufRcp1No);
				// @20131017
			}else if (iMode == 2){
				//選択レシピＮｏ（比較先）取得 @20131017
				sRcp2Info = (String)cmbRcp2.getSelectedItem();
				if (sRcp2Info != null)
				{
					if (sRcp2Info.indexOf(" ") != -1)
						bufRcp2No = Integer.valueOf(sRcp2Info.substring(0,sRcp2Info.indexOf(" "))).intValue();
				}
	            CZSystem.log("CZTblHikaku2 ChgRcp2No","（比較先）レシピ番号変更");
	            CZSystem.log("CZTblHikaku2 ChgRcp2No","（比較先）レシピ番号取得："+ bufRcp2No);
				// @20131017
			}

			if (iMode == 1)			//from
				cmbRcp1.removeAllItems();
			else if (iMode == 2)	//to
				cmbRcp2.removeAllItems();
			
			//選択された、炉番とテーブルの情報を格納する。
			
			Vector p_list = CZSystem.getCtTbRcp(roName, g_no);
			if (p_list != null)
			{
				for(int i = 0 ; p_list.size() > i ; i++){
					dtTile = (CZSystemCtTitle)p_list.elementAt(i);
					if (iMode == 1)			//from
					{
						cmbRcp1.addItem( dtTile.r_no + " : " + dtTile.title.trim());
						
						// @20131017 前回選択したレシピNoが変更炉番先に同じレシピNoがあるか？
						int icnt = cmbRcp1.getItemCount();
						CZSystem.log("CZTblHikaku2 setRcpDt","コンボボックス内 レシピ数：" + icnt);
						CZSystem.log("CZTblHikaku2 setRcpDt","コンボボックス内 レシピNo：" + dtTile.r_no);
						CZSystem.log("CZTblHikaku2 setRcpDt","コンボボックス内 前回選択レシピNo：" + bufRcp1No);
						if (dtTile.r_no == bufRcp1No)  // 同じレシピNoがある場合のindex取得
						{
							index1 = icnt;
							CZSystem.log("CZTblHikaku2 setRcpDt","コンボボックス Index1数：" + index1);
						}
						// @20131017
					}
					else if (iMode == 2)	//to
					{
						cmbRcp2.addItem( dtTile.r_no + " : " +  dtTile.title.trim());
						
						// @20131017 前回選択したレシピNoが変更炉番元に同じレシピNoがあるか？
						int icnt = cmbRcp2.getItemCount();
						CZSystem.log("CZTblHikaku2 setRcpDt","コンボボックス内 レシピ数：" + icnt);
						CZSystem.log("CZTblHikaku2 setRcpDt","コンボボックス内 レシピNo：" + dtTile.r_no);
						CZSystem.log("CZTblHikaku2 setRcpDt","コンボボックス内 前回選択レシピNo：" + bufRcp2No);
						if (dtTile.r_no == bufRcp2No)  // 同じレシピNoがある場合のindex取得
						{
							index2 = icnt;
							CZSystem.log("CZTblHikaku2 setRcpDt","コンボボックス Index2数：" + index2);
						}
						// @20131017
					}
				}
				//@20131017 前回選択レシピ番号を指定
				CZSystem.log("CZTblHikaku2 setRcpDt","前回選択レシピ番号indexを指定 : " + index1);
				if (index1 != -1)
				{
					cmbRcp1.setSelectedIndex(index1-1);
				}
				// @20131017
				//@20131017 前回選択レシピ番号を指定
				CZSystem.log("CZTblHikaku2 setRcpDt","前回選択レシピ番号index2を指定 : " + index2);
				if (index2 != -1)
				{
					cmbRcp2.setSelectedIndex(index2-1);
				}
				// @20131017
			}
		}
	}

	//
	// 制御テーブルのグループNoを格納する
    public class GrupNo extends JComboBox {
		String	sData[] = {"Ｔ１","Ｔ２","Ｔ３","Ｔ４","Ｔ５","Ｔ６"};
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
			/* T1からT6を返す */
			return((String)getSelectedItem());
		}

		public int getGrupNo() {
			/* T1からT6を1から6で返す */
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

				CZSystem.log("ChgRoNo", "炉番１ : " + roName1 + "炉番２：" + roName2 + "g_no : " + g_no);
				
				//選択レシピＮｏ（比較元）取得 @20131017
				sRcp1Info = (String)cmbRcp1.getSelectedItem();
				if (sRcp1Info != null)
				{
					if (sRcp1Info.indexOf(" ") != -1)
						bufRcp1No = Integer.valueOf(sRcp1Info.substring(0,sRcp1Info.indexOf(" "))).intValue();
				}
	            CZSystem.log("CZTblHikaku2 ChgRcp1No","（比較元）レシピ番号変更");
	            CZSystem.log("CZTblHikaku2 ChgRcp1No","（比較元）レシピ番号取得："+ bufRcp1No);
				// @20131017

				//選択レシピＮｏ（比較先）取得 @20131017
				sRcp2Info = (String)cmbRcp2.getSelectedItem();
				if (sRcp2Info != null)
				{
					if (sRcp2Info.indexOf(" ") != -1)
						bufRcp2No = Integer.valueOf(sRcp2Info.substring(0,sRcp2Info.indexOf(" "))).intValue();
				}
	            CZSystem.log("CZTblHikaku2 ChgRcp2No","（比較先）レシピ番号変更");
	            CZSystem.log("CZTblHikaku2 ChgRcp2No","（比較先）レシピ番号取得："+ bufRcp2No);
				// @20131017

				cmbRcp1.removeAllItems();
				cmbRcp2.removeAllItems();
				
				//選択された、炉番とテーブルの情報を格納する。
				Vector p_list = CZSystem.getCtTbRcp(roName1, g_no);
				if (p_list != null)
				{
					for(int i = 0 ; p_list.size() > i ; i++){
						dtTile = (CZSystemCtTitle)p_list.elementAt(i);
						cmbRcp1.addItem( dtTile.r_no + " : " + dtTile.title.trim());
						
						// @20131017 前回選択したレシピNoが変更炉番先に同じレシピNoがあるか？
						int icnt = cmbRcp1.getItemCount();
						CZSystem.log("CZTblHikaku2 setRcpDt","コンボボックス内 レシピ数：" + icnt);
						CZSystem.log("CZTblHikaku2 setRcpDt","コンボボックス内 レシピNo：" + dtTile.r_no);
						CZSystem.log("CZTblHikaku2 setRcpDt","コンボボックス内 前回選択レシピNo：" + bufRcp1No);
						if (dtTile.r_no == bufRcp1No)  // 同じレシピNoがある場合のindex取得
						{
							index1 = icnt;
							CZSystem.log("CZTblHikaku2 setRcpDt","コンボボックス Index数：" + index1);
						}
						// @20131017
					
					}
					//@20131017 前回選択レシピ番号を指定
					CZSystem.log("CZTblHikaku2 setRcpDt","前回選択レシピ番号index1を指定 : " + index1);
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
						
						// @20131017 前回選択したレシピNoが変更炉番先に同じレシピNoがあるか？
						int icnt = cmbRcp2.getItemCount();
						CZSystem.log("CZTblHikaku2 setRcpDt","コンボボックス内 レシピ数：" + icnt);
						CZSystem.log("CZTblHikaku2 setRcpDt","コンボボックス内 レシピNo：" + dtTile.r_no);
						CZSystem.log("CZTblHikaku2 setRcpDt","コンボボックス内 前回選択レシピNo：" + bufRcp2No);
						if (dtTile.r_no == bufRcp2No)  // 同じレシピNoがある場合のindex取得
						{
							index2 = icnt;
							CZSystem.log("CZTblHikaku2 setRcpDt","コンボボックス Index2数：" + index2);
						}
						// @20131017
					
					}
					//@20131017 前回選択レシピ番号を指定
					CZSystem.log("CZTblHikaku2 setRcpDt","前回選択レシピ番号index2を指定 : " + index2);
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
