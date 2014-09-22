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
 *   制御テーブル比較用Window 
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
	
	String sHikakuHed = ",#,項目,Min,Max,桁,単位,値,値,";
	String sLine   = new String("");
	String sDtOut  = new String("");
	
	//
	//
	//
	CZTblRcpOutPut(){
		super();
		
		setTitle("出力レシピ選択");
		
		setSize(460,230);
		setResizable(false);
		setModal(true);
		
		getContentPane().setLayout(null);
        // 他基地参照機能    @20131021
        if(CZSystemDefine.REFERENCE_RUN == CZSystem.getRunLevel()){
            getContentPane().setBackground(CZSystemDefine.DEFAULT_REFERENCE_BACKGROUND_COL);
        }else{
            getContentPane().setBackground(CZSystemDefine.DEFAULT_BACKGROUND_COL);
        }
		
		
		JLabel  lab = new JLabel("炉番",JLabel.CENTER);
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
		
		robutton = new JButton("▼");
		robutton.setBounds(220, 20, 30, 30);
		robutton.setLocale(new Locale("ja","JP"));
		robutton.setFont(new java.awt.Font("dialog", 0, 18));
		robutton.setBorder(new Flush3DBorder());
		robutton.setBackground(java.awt.Color.lightGray);  
		robutton.addActionListener(this);
		getContentPane().add(robutton);
		
		lab = new JLabel("グループ",JLabel.CENTER);
		lab.setBounds(20, 70, 100, 30);
		lab.setLocale(new Locale("ja","JP"));
		lab.setFont(new java.awt.Font("dialog", 0, 18));
		lab.setBorder(new Flush3DBorder());
		lab.setForeground(java.awt.Color.black);
		getContentPane().add(lab);
		
		cmdGrup = new GrupNo();
		cmdGrup.setBounds(120, 70, 100, 30);
		getContentPane().add(cmdGrup);
		
		
		lab = new JLabel("レシピNo",JLabel.CENTER);
		lab.setBounds(20, 100, 100, 30);
		lab.setLocale(new Locale("ja","JP"));
		lab.setFont(new java.awt.Font("dialog", 0, 18));
		lab.setBorder(new Flush3DBorder());
		lab.setForeground(java.awt.Color.black);
		getContentPane().add(lab);
		
		cmbRcp1 = new JComboBox();
		cmbRcp1.setBounds(120, 100, 300, 30);
		//ＴＯ炉選択中レシピ格納
		/*ro_from.setRcpDt();*/
		getContentPane().add(cmbRcp1);
		
		hikaku_btn = new JButton("出　力");
		hikaku_btn.setBounds(20, 150, 100, 24);
		hikaku_btn.setLocale(new Locale("ja","JP"));
		hikaku_btn.setFont(new java.awt.Font("dialog", 0, 18));
		hikaku_btn.setBorder(new Flush3DBorder());
		hikaku_btn.setForeground(java.awt.Color.black);
		hikaku_btn.addActionListener(new hikaku_btn_click());
		getContentPane().add(hikaku_btn);
		
		cancel_button = new JButton("終  了");
		cancel_button.setBounds(140, 150, 100, 24);
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
	* アクションリスナー
	* @param event 
	* @return none
	*/
	
	public void actionPerformed( ActionEvent event ) {
		Object source = event.getSource();
	
		CZRoSelectWin2 rosel;
		
		if( source == robutton){
			CZSystem.log("CZTbleRcpOutPut","イベントゲット！！！");
			
			rosel = new CZRoSelectWin2();
			rosel.setVisible(true);
			
			if(RoNameField.getText() != null){
				String roName;
				int g_no;
				CZSystemCtTitle   dtTile;
				roName = RoNameField.getText();
				
				g_no = cmdGrup.getGrupNo();
				
				CZSystem.log("setRcpDt", "炉番 : " + roName + "g_no : " + g_no );
				
				cmbRcp1.removeAllItems();
				
				//選択された、炉番とテーブルの情報を格納する。
				if(0 != CZSystemDefine.DISP_KETA_FLG){		/* 佐賀200mm用 */
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
			
			int tblGno = cmdGrup.getGrupNo();				//選択グループＮｏ取得
			String from_ro = RoNameField.getText();	//選択炉名取得
			String sBuf = null;								//変換用仮バッファー
			int rcp_from = 0;								//選択レシピ
			
			if(0 != CZSystemDefine.DISP_KETA_FLG){		/* 佐賀200mm用 */
				StringBuffer a = new StringBuffer();
				a.append(from_ro);
				a.insert(0,"K");
				from_ro = a.toString();
			}
			
			//選択レシピＮｏ取得
			sBuf = (String)cmbRcp1.getSelectedItem();
			if (sBuf != null)
			{
				if (sBuf.indexOf(" ") != -1)
					rcp_from = Integer.valueOf(sBuf.substring(0,sBuf.indexOf(" "))).intValue();
			}
			else
			{
				JOptionPane.showMessageDialog(null,"出力レシピデータなし","レシピ出力",JOptionPane.ERROR_MESSAGE);
//errorMsg("比較処理", "比較元のレシピNoなし");
				CZSystem.log("hikaku_btn_click","レシピなし");
				return;
			}
			
			//選択グループ判定
			if (tblGno == 6)
			{
				//比較処理実施
				subT6Chk(RoNameField.getText(), from_ro, rcp_from);
			}
			else
			{
				//比較処理実施
				subT1_5Chk(tblGno, RoNameField.getText(), from_ro, rcp_from);
			}
		}
	}
	
	private void subT1_5Chk(int tblGno, String fromDB_ro, String from_ro, int rcp_from){
//		int		tblGno;		選択されたグループNo
//		String fromDB_ro;	選択炉名（比較元のＤＢ名称）
//		String from_ro;		選択炉名（比較元の表示名称）
//		int rcp_from;		選択レシピ（比較元）
//		String toDB_ro;		選択炉名（比較先のＤＢ名称）
//		String to_ro;		選択炉名（比較先の表示名称）
//		int rcp_to;			選択レシピ（比較先）
		
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
		
		
		String sBuf = null;								//変換用仮バッファー
		
		//名称取得
		dataName = CZSystem.ctTblAllNameRead(tblGno);
		
		//比較元のレシピ情報取得
		data = CZSystem.getCtAllTb(from_ro, tblGno, rcp_from);
		
		//現在時刻取得
		String sNowDate = CZSystem.getDateTime("yyyy/MM/dd HH:mm:ss");
		
		//ファイル生成
		File file = new File(CZSystem.RECIPE_OUTPUT_PATH, "制御レシピ出力" + from_ro + "-" + "T" + tblGno + "-" + rcp_from + "_" + CZSystem.getDateTime("yyMMddHHmm") + ".csv");
		PrintWriter pr     = null;
		FileOutputStream s = null;
		
		try{
			String rs = CZSystem.RoKetaChg(from_ro);
			
			s = new FileOutputStream(file);
			pr = new PrintWriter(s);
			
			sLine = "制御テーブルデータ出力（ T" + tblGno + " ）,,,,,,,,," + sNowDate;
			pr.println(sLine);
			sLine = "";
			pr.println(sLine);
			sLine = "★★★★★★★出力レシピ情報★★★★★★★";
			pr.println(sLine);
			sLine = "";
			pr.println(sLine);
			sLine = "炉No1," + rs + ",";
			pr.println(sLine);
			sLine = "レシピNo," + rcp_from + ",";
			pr.println(sLine);
			sLine = "";
			pr.println(sLine);
			
			sLine = "★★★★★★★レシピ内容★★★★★★★";
			pr.println(sLine);
			
			if ((null != dataName) &&(null != data)){
				iMax1 = data.size();
				iRec1 = 0;
				iNameRec = 0;
				
				for(int i = 0; ((iNameRec < dataName.size()) || (iRec1 < iMax1)); i++){
					/* データチェック */
					if (iNameRec < dataName.size()){
						dtName = (CZSystemCtName)dataName.elementAt(iNameRec);
					}
					
					if (iRec1 < iMax1){
						iChkDtRtc = 11;
						d = (CZSystemCtTb)data.elementAt(iRec1);
					}
					else
					{
						/* 見出しのみ出力 */
						iChkDtRtc = -99;
					}
					
					sDtOut = ",";	/* データ部クリア */
					iRt = -1;
					
					/* ******* 見出しチェック ******** */
					if (iNameRec >= dataName.size()){
						iRt = 2;	/* 見出しなし */
					}
					else if (iRec1 >= iMax1){
						iRt = 1;	/* 未設定項目出力(データ抜けの項目出力) */
					}
					else
					{
						iRt = chkDtName(dtName, d);
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
							sDtOut = ",Ｌ軸,Ｒ軸,";
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
						sDtOut = ",Ｌ軸,Ｒ軸,";
						pr.println(sDtOut);
						iNameRec++;	/* 項目名称レコードチェンジ */
					}
					else if (iRt == 2)
					{	/* 該当見出しなし */
						if (iMidasi1 != d.t_no)
						{
							sLine = "";
							pr.println(sLine);
							sLine = "【" + d.t_no + "】";
							pr.println(sLine);
							
							iMidasi1 = d.t_no;
						}
						sDtOut = ",Ｌ軸,Ｒ軸,";
						pr.println(sDtOut);
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
								break;	//ターゲット変更
							}
							
							if (iChkDtRtc == 11)
							{	/* データ違う（データ１しかない） */
								sDtOut = ",";
								sDtOut += d.l_val + "," + d.r_val + ",";
								iRec1++;	/* データ１レコードUP */
								
								pr.println(sDtOut);
							}
						}
					}
				}	/* For End */
				JOptionPane.showMessageDialog(null,"出力が完了しました。","出力処理",JOptionPane.INFORMATION_MESSAGE);
			}
			else
			{
				if (null == dataName)
				{
					sLine = "";
					pr.println(sLine);
					sLine = "制御テーブル　項目定義がありません";
					pr.println(sLine);
				}
				
				if (null == data)
				{
					sLine = "";
					pr.println(sLine);
					sLine = "制御テーブルのデータがありません";
					pr.println(sLine);
				}
			}
		}
		catch(IOException e){
			if(null != pr) pr.close();
		}
		
		if(null != pr) pr.close();
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
	
	 private void subT6Chk(String fromDB_ro, String from_ro, int rcp_from){
//		String fromDB_ro;	選択炉名（比較元のＤＢ名称）
//		String from_ro;		選択炉名（比較元の表示名称）
//		int rcp_from;		選択レシピ（比較元）
//		String toDB_ro;		選択炉名（比較先のＤＢ名称）
//		String to_ro;		選択炉名（比較先の表示名称）
//		int rcp_to;			選択レシピ（比較先）
		
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
		
		
		String sBuf = null;								//変換用仮バッファー
		
		//Ｔ６名称取得
		dataName = CZSystem.ctT6AllNameRead();
		
		//比較元のＴ６レシピ情報取得
		data = CZSystem.getCtT6AllTb(from_ro,rcp_from);
		
		//現在時刻取得
		String sNowDate = CZSystem.getDateTime("yyyy/MM/dd HH:mm:ss");
		
		//ファイル生成
		File file = new File(CZSystem.RECIPE_OUTPUT_PATH, "制御レシピ出力" + from_ro + "-" + "T6" + "-" + rcp_from +"_" + CZSystem.getDateTime("yyMMddHHmm") + ".csv");
		PrintWriter pr     = null;
		FileOutputStream s = null;
		
		
		try{
			String rs = CZSystem.RoKetaChg(from_ro);
			
			s = new FileOutputStream(file);
			pr = new PrintWriter(s);
			
			sLine = "制御テーブルデータ出力（T6）,,,,,,,,," + sNowDate;
			pr.println(sLine);
			sLine = "";
			pr.println(sLine);
			sLine = "★★★★★★★出力レシピ情報★★★★★★★";
			pr.println(sLine);
			sLine = "";
			pr.println(sLine);
			sLine = "炉No1," + rs + ",";
			pr.println(sLine);
			sLine = "レシピNo," + rcp_from + ",";
			pr.println(sLine);
			sLine = "";
			pr.println(sLine);
			
			sLine = "★★★★★★★レシピ内容★★★★★★★";
			pr.println(sLine);
			
			if ((null != dataName) &&(null != data))
			{
				iMax1 = data.size();
				iRec1 = 0;
				iNameRec = 0;
				for(int i = 0; ((iNameRec < dataName.size()) || (iRec1 < iMax1)); i++){ 
					/* データチェック */
					if (iNameRec < dataName.size())
					{
						dtName = (CZSystemCtT6AllName)dataName.elementAt(iNameRec);
					}
					
					if (iRec1 < iMax1)
					{
						d = (CZSystemCtT6Tb)data.elementAt(iRec1);
					}
					
					
					/* ******* データ比較 ******** */
					if (iRec1 >= iMax1)
						iChkDtRtc = -99;
					else
						iChkDtRtc = 11;
					
					sDtOut = ",";	/* データ部クリア */
					iRt = -1;
					/* ******* 見出しチェック ******** */
					if (iNameRec >= dataName.size())
					{
						iRt = 2;	/* 見出しなし */
					}
					else if (iRec1 >= iMax1)
					{
						iRt = 1;	/* 未設定項目出力(データ抜けの項目出力) */
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
						if (iChkDtRtc == 11)
						{	/* データ違う（データ１しかない） */
							sDtOut += d.k_val + ",";
							pr.println(sDtOut);
							iRec1++;	/* データ１レコードUP */
						}
					}
				}	/* for end */
				//infoMsg("出力処理","出力完了");
				JOptionPane.showMessageDialog(null,"出力が完了しました。","出力処理",JOptionPane.INFORMATION_MESSAGE);
			}
			else
			{
				if (null == data)
				{
					sLine = "";
					pr.println(sLine);
					sLine = "T6データなし";
					pr.println(sLine);
				}
				
				if (null == dataName)
				{
					sLine = "";
					pr.println(sLine);
					sLine = "T6定義データなし";
					pr.println(sLine);
				}
			}
		}
		catch(IOException e){
			if(null != pr) pr.close();
		}
		
		if(null != pr) pr.close();
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
//			setDefault();
			setVisible(false);
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
				
				String roName1 = RoNameField.getText();
				int g_no = cmdGrup.getGrupNo();
				CZSystemCtTitle dtTile = null;
				
				cmbRcp1.removeAllItems();
				
				//選択された、炉番とテーブルの情報を格納する。
				if(0 != CZSystemDefine.DISP_KETA_FLG){		/* 佐賀200mm用 */
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
