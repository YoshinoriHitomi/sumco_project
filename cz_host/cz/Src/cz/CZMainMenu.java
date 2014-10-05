package cz;

import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.util.Locale;

import javax.swing.JMenu;
import javax.swing.JMenuBar;
import javax.swing.JMenuItem;

/***********************************************************
 *
 *   メイン画面用メニューバー
 * @author  (KPK Co.,Ltd.)
 * @version 1.0 (2003/04/01)
 * 2008.09.29 H.Nagamine 操業定数・制御テーブル変更履歴作成
 ***********************************************************/
public class CZMainMenu extends JMenuBar {

    // ----- コンストラクタ --------------------------------
    //
    CZMainMenu(){
        super();

        setLocale(new Locale("ja","JP"));
        setFont(new java.awt.Font("dialog", 0, 18));

        JMenu table         = null;
        JMenu record        = null;
        JMenu measurement   = null;
        JMenu error         = null;
        JMenu mainte        = null;
        JMenu information   = null;
        JMenu cms           = null;

        table = new JMenu("テーブル");
        setTable(table);
        add(table);

        record = new JMenu("実績");
        setRecord(record);
        add(record);

        measurement = new JMenu("計測");
        setMeasurement(measurement);
        add(measurement);

        error = new JMenu("エラー");
        setError(error);
        add(error);

        mainte = new JMenu("メンテナンス");
        setMainte(mainte);
        add(mainte);

        information = new JMenu("情報");
        setInformation(information);
        add(information);

/********************* 2007.08.29 cut **************************
        if(CZSystemDefine.ADMIN_RUN == CZSystem.getRunLevel()){
            cms  = new JMenu("集中監視");
            setCMS(cms);
            add(cms);

        }
********************* 2007.08.29 cut **************************/

    }

    // ----- ここからメニューアイテム ----------------------
    //
    //テーブルのサブメニュー
    //
    private void setTable(JMenu obj){

        try{
            JMenuItem item;

            item  = new JMenuItem("操業定数設定");
            item.setLocale(new Locale("ja","JP"));
            item.setFont(new java.awt.Font("dialog", 0, 18));
            item.addActionListener(new OperationTable());
            obj.add(item);

            obj.addSeparator();
            item  = new JMenuItem("制御テーブル");
            item.setLocale(new Locale("ja","JP"));
            item.setFont(new java.awt.Font("dialog", 0, 18));
            item.addActionListener(new ControlTable());
            obj.add(item);

//2006.06.02　y.k
            obj.addSeparator();
            obj.addSeparator();

            item  = new JMenuItem("操業定数　比較");
            item.setLocale(new Locale("ja","JP"));
            item.setFont(new java.awt.Font("dialog", 0, 18));
            item.addActionListener(new TblHikaku1());
            obj.add(item);

            item  = new JMenuItem("制御テーブル　比較");
            item.setLocale(new Locale("ja","JP"));
            item.setFont(new java.awt.Font("dialog", 0, 18));
            item.addActionListener(new TblHikaku2());
            obj.add(item);

//2006.06.02　y.k　end

            obj.addSeparator();
            obj.addSeparator();

            item  = new JMenuItem("レシピ内容出力");
            item.setLocale(new Locale("ja","JP"));
            item.setFont(new java.awt.Font("dialog", 0, 18));
            item.addActionListener(new TblRcpOutPut());
            obj.add(item);

            obj.addSeparator();
            obj.addSeparator();
// add start 2008.09.29
            item  = new JMenuItem("変更履歴出力");
            item.setLocale(new Locale("ja","JP"));
            item.setFont(new java.awt.Font("dialog", 0, 18));
            item.addActionListener(new ModifyOutPut());
            obj.add(item);

            obj.addSeparator();
            obj.addSeparator();
// add end 2008.09.29

            item  = new JMenuItem("終り");
            item.setLocale(new Locale("ja","JP"));
            item.setFont(new java.awt.Font("dialog", 0, 18));
            item.addActionListener(new Exit());
            obj.add(item);

        }
        catch (Throwable e) {
            CZSystem.handleException(e);
        }
    return ;
    }

    // -----------------------------------------------------
    //実績のサブメニュー
    //
    private void setRecord(JMenu obj){

        try{
            JMenuItem item;

            item  = new JMenuItem("ＰＶ保存");
            item.setLocale(new Locale("ja","JP"));
            item.setFont(new java.awt.Font("dialog", 0, 18));
            item.addActionListener(new PVData());
            obj.add(item);

//            item  = new JMenuItem("ＴＰＧ");
//            item.setLocale(new Locale("ja","JP"));
//            item.setFont(new java.awt.Font("dialog", 0, 18));
//            item.addActionListener(new TPGMain());
//            obj.add(item);

            item  = new JMenuItem("ＴＰＧ");
//              item  = new JMenuItem("トレンドテーブル");
            item.setLocale(new Locale("ja","JP"));
            item.setFont(new java.awt.Font("dialog", 0, 18));
            item.addActionListener(new TPGMain2());
            obj.add(item);

            item  = new JMenuItem("ＦｐＡｖｅ");
            item.setLocale(new Locale("ja","JP"));
            item.setFont(new java.awt.Font("dialog", 0, 18));
            item.addActionListener(new FpAveMain());
            obj.add(item);

            item  = new JMenuItem("複数ＰＶ");
            item.setLocale(new Locale("ja","JP"));
            item.setFont(new java.awt.Font("dialog", 0, 18));
            item.addActionListener(new PVSomeData());
            obj.add(item);
        }
        catch (Throwable e) {
            CZSystem.handleException(e);
        }
        return ;
    }

    // -----------------------------------------------------
    // 計測のサブメニュー
    //
    private void setMeasurement(JMenu obj){

        try{
            JMenuItem item;

            item  = new JMenuItem("ＣＣＤ生波形");
            item.setLocale(new Locale("ja","JP"));
            item.setFont(new java.awt.Font("dialog", 0, 18));
            item.addActionListener(new CMSCCDWave());
            obj.add(item);

            item  = new JMenuItem("ＣＣＤ画像");
            item.setLocale(new Locale("ja","JP"));
            item.setFont(new java.awt.Font("dialog", 0, 18));
            item.addActionListener(new CMSCCDBMP());
            obj.add(item);

            item  = new JMenuItem("輝度変化チェック");
            item.setLocale(new Locale("ja","JP"));
            item.setFont(new java.awt.Font("dialog", 0, 18));
            item.addActionListener(new CMSBrightnessCheck());
            obj.add(item);

        }
        catch (Throwable e) {
            CZSystem.handleException(e);
        }
        return ;
    }

    // -----------------------------------------------------
    //エラーのサブメニュー
    //
    private void setError(JMenu obj){

        try{
            JMenuItem item;

            item  = new JMenuItem("システムエラー");
            item.setLocale(new Locale("ja","JP"));
            item.setFont(new java.awt.Font("dialog", 0, 18));
            item.addActionListener(new ErrorMsgWin());
            obj.add(item);

            obj.addSeparator();

            item  = new JMenuItem("サーバーシステムエラー");
            item.setLocale(new Locale("ja","JP"));
            item.setFont(new java.awt.Font("dialog", 0, 18));
            item.addActionListener(new HostErrorMsgWin());
            obj.add(item);

        }
        catch (Throwable e) {
            CZSystem.handleException(e);
        }
        return ;
    }

    // -----------------------------------------------------
    //メンテナンスのサブメニュー
    //
    private void setMainte(JMenu obj){

        try{
            JMenuItem item;

            item  = new JMenuItem("エラー項目定義");
            item.setLocale(new Locale("ja","JP"));
            item.setFont(new java.awt.Font("dialog", 0, 18));
            item.addActionListener(new ErrorSetWin());
            obj.add(item);

            obj.addSeparator();
            item  = new JMenuItem("操業定数コピー");
            item.setLocale(new Locale("ja","JP"));
            item.setFont(new java.awt.Font("dialog", 0, 18));
            item.addActionListener(new OperationTableCp());
            obj.add(item);
/*@@ ＭＯ管理
            obj.addSeparator();
            item  = new JMenuItem("ＭＯ管理");
            item.setLocale(new Locale("ja","JP"));
            item.setFont(new java.awt.Font("dialog", 0, 18));
            item.addActionListener(new MOControl());
            obj.add(item);
@@*/
/*
            obj.addSeparator();
            item  = new JMenuItem("ＲＡＩＤ状態");
            item.setLocale(new Locale("ja","JP"));
            item.setFont(new java.awt.Font("dialog", 0, 18));
            item.addActionListener(new RaidWatch());
            obj.add(item);
*/
        }
        catch (Throwable e) {
            CZSystem.handleException(e);
        }
        return ;
    }


    // -----------------------------------------------------
    // 情報のサブメニュー
    //
    private void setInformation(JMenu obj){

        try{
            JMenuItem item;

            item  = new JMenuItem("稼働状況");
            item.setLocale(new Locale("ja","JP"));
            item.setFont(new java.awt.Font("dialog", 0, 18));
            item.addActionListener(new OperationalStatus());
            obj.add(item);

			/* 2006.07.10 start*/
            item  = new JMenuItem("監視状況");
            item.setLocale(new Locale("ja","JP"));
            item.setFont(new java.awt.Font("dialog", 0, 18));
            item.addActionListener(new AllRealStatus());
            obj.add(item);
			/* 2006.07.10 end */

			/* 2008.09.10 start*/
            item  = new JMenuItem("排他状況");
            item.setLocale(new Locale("ja","JP"));
            item.setFont(new java.awt.Font("dialog", 0, 18));
            item.addActionListener(new HaitaStatus());
            obj.add(item);
			/* 2008.09.10 end */

        }
        catch (Throwable e) {
            CZSystem.handleException(e);
        }
        return ;
    }


    // -----------------------------------------------------
    // 集中監視のサブメニュー
    //
    private void setCMS(JMenu obj){

        try{
            JMenuItem item;

            item  = new JMenuItem("電源");
            item.setLocale(new Locale("ja","JP"));
            item.setFont(new java.awt.Font("dialog", 0, 18));
            item.addActionListener(new CMSPower());
            obj.add(item);

            obj.addSeparator();
            item  = new JMenuItem("制御モード");
            item.setLocale(new Locale("ja","JP"));
            item.setFont(new java.awt.Font("dialog", 0, 18));
            item.addActionListener(new CMSControlMode());
            obj.add(item);

            obj.addSeparator();
            item  = new JMenuItem("プロセス");
            item.setLocale(new Locale("ja","JP"));
            item.setFont(new java.awt.Font("dialog", 0, 18));
            item.addActionListener(new CMSProcChg());
            obj.add(item);

            obj.addSeparator();
            item  = new JMenuItem("シード速度");
            item.setLocale(new Locale("ja","JP"));
            item.setFont(new java.awt.Font("dialog", 0, 18));
            item.addActionListener(new CMSSeedSpeed());
            obj.add(item);

            item  = new JMenuItem("シード回転");
            item.setLocale(new Locale("ja","JP"));
            item.setFont(new java.awt.Font("dialog", 0, 18));
            item.addActionListener(new CMSSeedRotation());
            obj.add(item);

            item  = new JMenuItem("シード位置");
            item.setLocale(new Locale("ja","JP"));
            item.setFont(new java.awt.Font("dialog", 0, 18));
            item.addActionListener(new CMSSeedPosition());
            obj.add(item);

            obj.addSeparator();
            item  = new JMenuItem("ルツボ速度");
            item.setLocale(new Locale("ja","JP"));
            item.setFont(new java.awt.Font("dialog", 0, 18));
            item.addActionListener(new CMSCrucibleSpeed());
            obj.add(item);

            item  = new JMenuItem("ルツボ回転");
            item.setLocale(new Locale("ja","JP"));
            item.setFont(new java.awt.Font("dialog", 0, 18));
            item.addActionListener(new CMSCrucibleRotation());
            obj.add(item);

            item  = new JMenuItem("ルツボ位置");
            item.setLocale(new Locale("ja","JP"));
            item.setFont(new java.awt.Font("dialog", 0, 18));
            item.addActionListener(new CMSCruciblePosition());
            obj.add(item);

            obj.addSeparator();
            item  = new JMenuItem("結晶保持");
            item.setLocale(new Locale("ja","JP"));
            item.setFont(new java.awt.Font("dialog", 0, 18));
            item.addActionListener(new CMSSXLHoldPosition());
            obj.add(item);

            obj.addSeparator();
            item  = new JMenuItem("メインヒータ１電力");
            item.setLocale(new Locale("ja","JP"));
            item.setFont(new java.awt.Font("dialog", 0, 18));
            item.addActionListener(new CMSHeater1Power());
            obj.add(item);

            item  = new JMenuItem("メインヒータ２電力");
            item.setLocale(new Locale("ja","JP"));
            item.setFont(new java.awt.Font("dialog", 0, 18));
            item.addActionListener(new CMSHeater2Power());
            obj.add(item);

            item  = new JMenuItem("ボトムヒータ電力");
            item.setLocale(new Locale("ja","JP"));
            item.setFont(new java.awt.Font("dialog", 0, 18));
            item.addActionListener(new CMSHeaterBottomPower());
            obj.add(item);

            item  = new JMenuItem("ヒータ温度");
            item.setLocale(new Locale("ja","JP"));
            item.setFont(new java.awt.Font("dialog", 0, 18));
            item.addActionListener(new CMSHeaterTemp());
            obj.add(item);

            item  = new JMenuItem("シードヒータ電力");
            item.setLocale(new Locale("ja","JP"));
            item.setFont(new java.awt.Font("dialog", 0, 18));
            item.addActionListener(new CMSHeaterSeedPower());
            obj.add(item);

            obj.addSeparator();
            item  = new JMenuItem("プルアルゴン");
            item.setLocale(new Locale("ja","JP"));
            item.setFont(new java.awt.Font("dialog", 0, 18));
            item.addActionListener(new CMSArPull());
            obj.add(item);

            item  = new JMenuItem("トップアルゴン");
            item.setLocale(new Locale("ja","JP"));
            item.setFont(new java.awt.Font("dialog", 0, 18));
            item.addActionListener(new CMSArTop());
            obj.add(item);
/*@@ 磁場強度
            obj.addSeparator();
            item  = new JMenuItem("磁場強度１");
            item.setLocale(new Locale("ja","JP"));
            item.setFont(new java.awt.Font("dialog", 0, 18));
            obj.add(item);

            item  = new JMenuItem("磁場強度２");
            item.setLocale(new Locale("ja","JP"));
            item.setFont(new java.awt.Font("dialog", 0, 18));
            obj.add(item);

            obj.addSeparator();
            item  = new JMenuItem("特定プロセス");
            item.setLocale(new Locale("ja","JP"));
            item.setFont(new java.awt.Font("dialog", 0, 18));
            obj.add(item);
@@*/
            obj.addSeparator();
            item  = new JMenuItem("モニター切替え");
            item.setLocale(new Locale("ja","JP"));
            item.setFont(new java.awt.Font("dialog", 0, 18));
            item.addActionListener(new CMSMonitorChg());
            obj.add(item);

        }
        catch (Throwable e) {
            CZSystem.handleException(e);
        }
        return ;
    }

    // ----- ここからActionListener ------------------------

    /*******************************************************
     *
     *   テーブル
     *
     *******************************************************/
    //
    // ----- 操業定数設定 ----------------------------------
    //
    class OperationTable implements ActionListener {

        private CZOperationTable obj = null;

        public void actionPerformed(ActionEvent e){
            CZSystem.getOperationMst();     //@@@@ 操業定数マスタを読込む
            if(null == obj) obj = new CZOperationTable();
            	if( 0 != CZSystemDefine.TIMER_FLG ){
            		obj.timerStart();
            	}
            obj.setDefault();
            obj.setVisible(true);
        }
    }

    // ----- 制御テーブル設定 ------------------------------
    //
    class ControlTable implements ActionListener {

        private CZControlTable obj = null;

        public void actionPerformed(ActionEvent e){
          CZSystem.getControlMst();     //@@@@ 制御テーブルマスタを読込む
          if(null == obj) obj = new CZControlTable();
          	if( 0 != CZSystemDefine.TIMER_FLG ){
	      		obj.timerStart();
	        }
          obj.setDefault();
          obj.setVisible(true);
        }
    }

    // ----- 操業テーブル比較 --2006.06.02 Y.K-------------
    //
    class TblHikaku1 implements ActionListener {

        private CZTblHikaku obj = null;

        public void actionPerformed(ActionEvent e){
          if(null == obj) obj = new CZTblHikaku(1);
//		  obj.setType(1);		//1:操業　2:制御
//          obj.setDefault();
          obj.setVisible(true);
        }
    }

    // ----- 制御テーブル比較 --2006.06.02 Y.K-------------
    //
    class TblHikaku2 implements ActionListener {

        private CZTblHikaku2 obj = null;

        public void actionPerformed(ActionEvent e){
          if(null == obj) obj = new CZTblHikaku2();
          obj.setVisible(true);
        }
    }

    // ----- レシピ内容出力 ---------------
    //
    class TblRcpOutPut implements ActionListener {

        private CZTblRcpOutPut obj = null;

        public void actionPerformed(ActionEvent e){
          if(null == obj) obj = new CZTblRcpOutPut();
          obj.setVisible(true);
        }
    }
// add start 2008.09.29
    // ----- 変更履歴出力 ---------------
    //
    class ModifyOutPut implements ActionListener {

        private CZModify obj = null;

        public void actionPerformed(ActionEvent e){
          if(null == obj) obj = new CZModify();
          obj.setDefault();
          obj.setVisible(true);
        }
    }
// add end 2008.09.29

    // ----- 終了 ------------------------------------------
    //
    class Exit implements ActionListener {
        public void actionPerformed(ActionEvent e){
            CZSystem.exit(0,"Menu Exit");
        }
    }

    /*******************************************************
     *
     *   実績
     *
     *******************************************************/
    //
    // ----- ＰＶ保存 --------------------------------------
    //
    class PVData implements ActionListener {
        private CZPVDataSave obj = null;

        public void actionPerformed(ActionEvent e){
            if(null == obj) obj = new CZPVDataSave();
            obj.setDefault();
            obj.setVisible(true);

        }
    }


    // ----- ＴＰＧ ----------------------------------------
    //
    class TPGMain implements ActionListener {
/*@@ CZTPGMain@@*/
        private CZTPGMain obj = null;

        public void actionPerformed(ActionEvent e){
            if(null == obj) obj = new CZTPGMain();
            obj.setDefault();
            obj.setVisible(true);

/*@@
        private CZTPGFrame obj = null;

        public void actionPerformed(ActionEvent e){
            if (null != obj) {
                obj = null;
            }
            obj = new CZTPGFrame();
            obj.setVisible(true);
@@*/
        }
    }

    // ----- トレンドテーブル ------------------------------
    //
    class TPGMain2 implements ActionListener {
/*@@ CZTPGMain
        private CZTPGMain obj = null;

        public void actionPerformed(ActionEvent e){
            if(null == obj) obj = new CZTPGMain();
            obj.setDefault();
            obj.setVisible(true);
@@*/

        private CZTPGFrame obj = null;

        public void actionPerformed(ActionEvent e){
            if (null != obj) {
                obj = null;
            }
            obj = new CZTPGFrame();
            obj.setVisible(true);
        }
    }

    // ----- ＦｐＡｖｅ ------------------------------------
    //
    class FpAveMain implements ActionListener {
        private CZFpAveMain obj = null;

        public void actionPerformed(ActionEvent e){
//@@            if(null == obj) obj = new CZFpAveMain();
            if (null != obj) {
                obj = null;
            }
            obj = new CZFpAveMain();
            obj.setDefault();
            obj.setVisible(true);
        }
    }

    //
    // ----- 複数ＰＶ保存 --------------------------------------
    //
    class PVSomeData implements ActionListener {
        private CZPVSomeDataSave obj = null;

        public void actionPerformed(ActionEvent e){
            if(null == obj) obj = new CZPVSomeDataSave();
            obj.setDefault();
            obj.setVisible(true);

        }
    }

    /*******************************************************
     *
     *   計測
     *
     *******************************************************/
    //
    // ----- ＣＣＤ生波形 ----------------------------------
    //
    class CMSCCDWave implements ActionListener {
        private CZCMSCCDWave obj = null;

        public void actionPerformed(ActionEvent e){
//          if(null == obj) obj = new CZCMSCCDWave();
            if(null != obj) obj = null;
            obj = new CZCMSCCDWave();
            obj.setDefault();
            obj.setVisible(true);
        }
    }

    // ----- ＣＣＤ画像 ------------------------------------
    //
    class CMSCCDBMP implements ActionListener {
        private CZCMSCCDBMP obj = null;

        public void actionPerformed(ActionEvent e){
            if(null == obj) obj = new CZCMSCCDBMP();
            obj.setDefault();
            obj.setVisible(true);
        }
    }

    //
    // ----- 輝度変化チェック ------------------------------
    //
    class CMSBrightnessCheck implements ActionListener {
        private CZCMSBrightnessCheck obj = null;

        public void actionPerformed(ActionEvent e){
//          if(null == obj) obj = new CZCMSBrightnessCheck();
            if(null != obj) obj = null;
            obj = new CZCMSBrightnessCheck();
            obj.setDefault();
            obj.setVisible(true);
        }
    }

    /*******************************************************
     *
     *   エラー
     *
     *******************************************************/
    //
    // ----- エラー表示 ------------------------------------
    //
    class ErrorMsgWin implements ActionListener {
        private CZErrorMsgWin obj = null;

        public void actionPerformed(ActionEvent e){
            if(null == obj) obj = new CZErrorMsgWin();
            obj.setDefault();
            obj.setVisible(true);
        }
    }

    // ----- サーバーエラー表示 ----------------------------
    //
    class HostErrorMsgWin implements ActionListener {
        private CZHostErrorMsgWin obj = null;

        public void actionPerformed(ActionEvent e){
            if(null == obj) obj = new CZHostErrorMsgWin();
            obj.setDefault();
            obj.setVisible(true);
        }
    }

    /*******************************************************
     *
     *   メンテナンス
     *
     *******************************************************/
    //
    // ----- エラー項目設定 --------------------------------
    //
    class ErrorSetWin implements ActionListener {
        private CZErrorSetWin obj = null;

        public void actionPerformed(ActionEvent e){
            if(null == obj) obj = new CZErrorSetWin();
            obj.setDefault();
            obj.setVisible(true);
        }
    }

    // ----- 操業定数コピー --------------------------------
    //
    class OperationTableCp implements ActionListener {
        private CZOperationTableCp obj = null;

        public void actionPerformed(ActionEvent e){
            if(null == obj) obj = new CZOperationTableCp();
            obj.setDefault();
            obj.setVisible(true);
        }
    }

    // ----- ＭＯ管理 --------------------------------------
    //
    class MOControl implements ActionListener {
        private CZMOControl obj = null;

        public void actionPerformed(ActionEvent e){
            if(null == obj) obj = new CZMOControl();
            obj.setDefault();
            obj.setVisible(true);
        }
    }

    // ----- ＲＡＩＤ状態 ----------------------------------
    //
    class RaidWatch implements ActionListener {
        private CZRaidWatch obj = null;

        public void actionPerformed(ActionEvent e){
            if(null == obj) obj = new CZRaidWatch();
            obj.setDefault();
            obj.setVisible(true);
        }
    }

    /*******************************************************
     *
     *   情報
     *
     *******************************************************/
    //
    // ----- 稼働状況 --------------------------------------
    //
    class OperationalStatus implements ActionListener {
        private CZOperationalStatus obj = null;

        public void actionPerformed(ActionEvent e){
            if(null == obj) obj = new CZOperationalStatus();
            obj.setDefault();
            obj.setVisible(true);
        }
    }

    //
    // ----- 監視状況 --------------------------------------
    // 2006.07.10
    class AllRealStatus implements ActionListener {
        private CZAllRealStatus obj = null;

        public void actionPerformed(ActionEvent e){
            if(null == obj) obj = new CZAllRealStatus();
            obj.setDefault();
            obj.setVisible(true);
        }
    }

    //
    // ----- 排他状況 --------------------------------------
    // 2008.09.10
    class HaitaStatus implements ActionListener {
        private CZHaitaStatus obj = null;

        public void actionPerformed(ActionEvent e){
            if(null == obj) obj = new CZHaitaStatus();
            obj.setDefault();
            obj.setVisible(true);
        }
    }

    /*******************************************************
     *
     *   集中監視
     *
     *******************************************************/
    //
    // ----- 電源 ------------------------------------------
    //
    class CMSPower implements ActionListener {
        private CZCMSPower obj = null;

        public void actionPerformed(ActionEvent e){
            if(null == obj) obj = new CZCMSPower();
            obj.setDefault();
            obj.setVisible(true);
        }
    }

    // ----- 制御モード ------------------------------------
    //
    class CMSControlMode implements ActionListener {
        private CZCMSControlMode obj = null;

        public void actionPerformed(ActionEvent e){
            if(null == obj) obj = new CZCMSControlMode();
            obj.setDefault();
            obj.setVisible(true);
        }
    }

    // ----- プロセス移行 ----------------------------------
    //
    class CMSProcChg implements ActionListener {
        private CZCMSProcChg obj = null;

        public void actionPerformed(ActionEvent e){
            if(null == obj) obj = new CZCMSProcChg();
            obj.setDefault();
            obj.setVisible(true);
        }
    }

    // ----- シード速度 ------------------------------------
    //
    class CMSSeedSpeed implements ActionListener {
        private CZCMSSeedSpeed obj = null;

        public void actionPerformed(ActionEvent e){
            if(null == obj) obj = new CZCMSSeedSpeed();
            obj.setDefault();
            obj.setVisible(true);
        }
    }

    // ----- シード回転 ------------------------------------
    //
    class CMSSeedRotation implements ActionListener {
        private CZCMSSeedRotation obj = null;

        public void actionPerformed(ActionEvent e){
            if(null == obj) obj = new CZCMSSeedRotation();
            obj.setDefault();
            obj.setVisible(true);
        }
    }

    // ----- シード位置 ------------------------------------
    //
    class CMSSeedPosition implements ActionListener {
        private CZCMSSeedPosition obj = null;

        public void actionPerformed(ActionEvent e){
            if(null == obj) obj = new CZCMSSeedPosition();
            obj.setDefault();
            obj.setVisible(true);
        }
    }

    // ----- ルツボ速度 ------------------------------------
    //
    class CMSCrucibleSpeed implements ActionListener {
        private CZCMSCrucibleSpeed obj = null;

        public void actionPerformed(ActionEvent e){
            if(null == obj) obj = new CZCMSCrucibleSpeed();
            obj.setDefault();
            obj.setVisible(true);
        }
    }

    // ----- ルツボ回転 ------------------------------------
    //
    class CMSCrucibleRotation implements ActionListener {
        private CZCMSCrucibleRotation obj = null;

        public void actionPerformed(ActionEvent e){
            if(null == obj) obj = new CZCMSCrucibleRotation();
            obj.setDefault();
            obj.setVisible(true);
        }
    }

    // ----- ルツボ位置 ------------------------------------
    //
    class CMSCruciblePosition implements ActionListener {
        private CZCMSCruciblePosition obj = null;

        public void actionPerformed(ActionEvent e){
            if(null == obj) obj = new CZCMSCruciblePosition();
            obj.setDefault();
            obj.setVisible(true);
        }
    }

    // ----- 結晶保持位置 ----------------------------------
    //
    class CMSSXLHoldPosition implements ActionListener {
        private CZCMSSXLHoldPosition obj = null;

        public void actionPerformed(ActionEvent e){
            if(null == obj) obj = new CZCMSSXLHoldPosition();
            obj.setDefault();
            obj.setVisible(true);
        }
    }

    // ----- ヒーター１パワー ------------------------------
    //
    class CMSHeater1Power implements ActionListener {
        private CZCMSHeater1Power obj = null;

        public void actionPerformed(ActionEvent e){
            if(null == obj) obj = new CZCMSHeater1Power();
            obj.setDefault();
            obj.setVisible(true);
        }
    }

    // ----- ヒーター２パワー ------------------------------
    //
    class CMSHeater2Power implements ActionListener {
        private CZCMSHeater2Power obj = null;

        public void actionPerformed(ActionEvent e){
            if(null == obj) obj = new CZCMSHeater2Power();
            obj.setDefault();
            obj.setVisible(true);
        }
    }

    // ----- ボトムヒーターパワー --------------------------
    //
    class CMSHeaterBottomPower implements ActionListener {
        private CZCMSHeaterBottomPower obj = null;

        public void actionPerformed(ActionEvent e){
            if(null == obj) obj = new CZCMSHeaterBottomPower();
            obj.setDefault();
            obj.setVisible(true);
        }
    }

    // ----- ヒーター１温度 --------------------------------
    //
    class CMSHeaterTemp implements ActionListener {
        private CZCMSHeaterTemp obj = null;

        public void actionPerformed(ActionEvent e){
            if(null == obj) obj = new CZCMSHeaterTemp();
            obj.setDefault();
            obj.setVisible(true);
        }
    }

    // ----- シードヒーター電力 ----------------------------
    //
    class CMSHeaterSeedPower implements ActionListener {
        private CZCMSHeaterSeedPower obj = null;

        public void actionPerformed(ActionEvent e){
            if(null == obj) obj = new CZCMSHeaterSeedPower();
            obj.setDefault();
            obj.setVisible(true);
        }
    }

    //
    // ----- プルアルゴン ----------------------------------
    //
    class CMSArPull implements ActionListener {
        private CZCMSArPull obj = null;

        public void actionPerformed(ActionEvent e){
            if(null == obj) obj = new CZCMSArPull();
            obj.setDefault();
            obj.setVisible(true);
        }
    }

    // ----- トップアルゴン --------------------------------
    //
    class CMSArTop implements ActionListener {
        private CZCMSArTop obj = null;

        public void actionPerformed(ActionEvent e){
            if(null == obj) obj = new CZCMSArTop();
            obj.setDefault();
            obj.setVisible(true);
        }
    }

    // ----- モニター切替え --------------------------------
    //
    class CMSMonitorChg implements ActionListener {
        private CZCMSMonitorChg obj = null;

        public void actionPerformed(ActionEvent e){
            if(null == obj) obj = new CZCMSMonitorChg();
            obj.setDefault();
            obj.setVisible(true);
        }
    }
}
