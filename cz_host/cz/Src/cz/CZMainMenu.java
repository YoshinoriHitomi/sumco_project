package cz;

import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.util.Locale;

import javax.swing.JMenu;
import javax.swing.JMenuBar;
import javax.swing.JMenuItem;

/***********************************************************
 *
 *   ���C����ʗp���j���[�o�[
 * @author  (KPK Co.,Ltd.)
 * @version 1.0 (2003/04/01)
 * 2008.09.29 H.Nagamine ���ƒ萔�E����e�[�u���ύX�����쐬
 ***********************************************************/
public class CZMainMenu extends JMenuBar {

    // ----- �R���X�g���N�^ --------------------------------
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

        table = new JMenu("�e�[�u��");
        setTable(table);
        add(table);

        record = new JMenu("����");
        setRecord(record);
        add(record);

        measurement = new JMenu("�v��");
        setMeasurement(measurement);
        add(measurement);

        error = new JMenu("�G���[");
        setError(error);
        add(error);

        mainte = new JMenu("�����e�i���X");
        setMainte(mainte);
        add(mainte);

        information = new JMenu("���");
        setInformation(information);
        add(information);

/********************* 2007.08.29 cut **************************
        if(CZSystemDefine.ADMIN_RUN == CZSystem.getRunLevel()){
            cms  = new JMenu("�W���Ď�");
            setCMS(cms);
            add(cms);

        }
********************* 2007.08.29 cut **************************/

    }

    // ----- �������烁�j���[�A�C�e�� ----------------------
    //
    //�e�[�u���̃T�u���j���[
    //
    private void setTable(JMenu obj){

        try{
            JMenuItem item;

            item  = new JMenuItem("���ƒ萔�ݒ�");
            item.setLocale(new Locale("ja","JP"));
            item.setFont(new java.awt.Font("dialog", 0, 18));
            item.addActionListener(new OperationTable());
            obj.add(item);

            obj.addSeparator();
            item  = new JMenuItem("����e�[�u��");
            item.setLocale(new Locale("ja","JP"));
            item.setFont(new java.awt.Font("dialog", 0, 18));
            item.addActionListener(new ControlTable());
            obj.add(item);

//2006.06.02�@y.k
            obj.addSeparator();
            obj.addSeparator();

            item  = new JMenuItem("���ƒ萔�@��r");
            item.setLocale(new Locale("ja","JP"));
            item.setFont(new java.awt.Font("dialog", 0, 18));
            item.addActionListener(new TblHikaku1());
            obj.add(item);

            item  = new JMenuItem("����e�[�u���@��r");
            item.setLocale(new Locale("ja","JP"));
            item.setFont(new java.awt.Font("dialog", 0, 18));
            item.addActionListener(new TblHikaku2());
            obj.add(item);

//2006.06.02�@y.k�@end

            obj.addSeparator();
            obj.addSeparator();

            item  = new JMenuItem("���V�s���e�o��");
            item.setLocale(new Locale("ja","JP"));
            item.setFont(new java.awt.Font("dialog", 0, 18));
            item.addActionListener(new TblRcpOutPut());
            obj.add(item);

            obj.addSeparator();
            obj.addSeparator();
// add start 2008.09.29
            item  = new JMenuItem("�ύX�����o��");
            item.setLocale(new Locale("ja","JP"));
            item.setFont(new java.awt.Font("dialog", 0, 18));
            item.addActionListener(new ModifyOutPut());
            obj.add(item);

            obj.addSeparator();
            obj.addSeparator();
// add end 2008.09.29

            item  = new JMenuItem("�I��");
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
    //���т̃T�u���j���[
    //
    private void setRecord(JMenu obj){

        try{
            JMenuItem item;

            item  = new JMenuItem("�o�u�ۑ�");
            item.setLocale(new Locale("ja","JP"));
            item.setFont(new java.awt.Font("dialog", 0, 18));
            item.addActionListener(new PVData());
            obj.add(item);

//            item  = new JMenuItem("�s�o�f");
//            item.setLocale(new Locale("ja","JP"));
//            item.setFont(new java.awt.Font("dialog", 0, 18));
//            item.addActionListener(new TPGMain());
//            obj.add(item);

            item  = new JMenuItem("�s�o�f");
//              item  = new JMenuItem("�g�����h�e�[�u��");
            item.setLocale(new Locale("ja","JP"));
            item.setFont(new java.awt.Font("dialog", 0, 18));
            item.addActionListener(new TPGMain2());
            obj.add(item);

            item  = new JMenuItem("�e���`����");
            item.setLocale(new Locale("ja","JP"));
            item.setFont(new java.awt.Font("dialog", 0, 18));
            item.addActionListener(new FpAveMain());
            obj.add(item);

            item  = new JMenuItem("�����o�u");
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
    // �v���̃T�u���j���[
    //
    private void setMeasurement(JMenu obj){

        try{
            JMenuItem item;

            item  = new JMenuItem("�b�b�c���g�`");
            item.setLocale(new Locale("ja","JP"));
            item.setFont(new java.awt.Font("dialog", 0, 18));
            item.addActionListener(new CMSCCDWave());
            obj.add(item);

            item  = new JMenuItem("�b�b�c�摜");
            item.setLocale(new Locale("ja","JP"));
            item.setFont(new java.awt.Font("dialog", 0, 18));
            item.addActionListener(new CMSCCDBMP());
            obj.add(item);

            item  = new JMenuItem("�P�x�ω��`�F�b�N");
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
    //�G���[�̃T�u���j���[
    //
    private void setError(JMenu obj){

        try{
            JMenuItem item;

            item  = new JMenuItem("�V�X�e���G���[");
            item.setLocale(new Locale("ja","JP"));
            item.setFont(new java.awt.Font("dialog", 0, 18));
            item.addActionListener(new ErrorMsgWin());
            obj.add(item);

            obj.addSeparator();

            item  = new JMenuItem("�T�[�o�[�V�X�e���G���[");
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
    //�����e�i���X�̃T�u���j���[
    //
    private void setMainte(JMenu obj){

        try{
            JMenuItem item;

            item  = new JMenuItem("�G���[���ڒ�`");
            item.setLocale(new Locale("ja","JP"));
            item.setFont(new java.awt.Font("dialog", 0, 18));
            item.addActionListener(new ErrorSetWin());
            obj.add(item);

            obj.addSeparator();
            item  = new JMenuItem("���ƒ萔�R�s�[");
            item.setLocale(new Locale("ja","JP"));
            item.setFont(new java.awt.Font("dialog", 0, 18));
            item.addActionListener(new OperationTableCp());
            obj.add(item);
/*@@ �l�n�Ǘ�
            obj.addSeparator();
            item  = new JMenuItem("�l�n�Ǘ�");
            item.setLocale(new Locale("ja","JP"));
            item.setFont(new java.awt.Font("dialog", 0, 18));
            item.addActionListener(new MOControl());
            obj.add(item);
@@*/
/*
            obj.addSeparator();
            item  = new JMenuItem("�q�`�h�c���");
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
    // ���̃T�u���j���[
    //
    private void setInformation(JMenu obj){

        try{
            JMenuItem item;

            item  = new JMenuItem("�ғ���");
            item.setLocale(new Locale("ja","JP"));
            item.setFont(new java.awt.Font("dialog", 0, 18));
            item.addActionListener(new OperationalStatus());
            obj.add(item);

			/* 2006.07.10 start*/
            item  = new JMenuItem("�Ď���");
            item.setLocale(new Locale("ja","JP"));
            item.setFont(new java.awt.Font("dialog", 0, 18));
            item.addActionListener(new AllRealStatus());
            obj.add(item);
			/* 2006.07.10 end */

			/* 2008.09.10 start*/
            item  = new JMenuItem("�r����");
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
    // �W���Ď��̃T�u���j���[
    //
    private void setCMS(JMenu obj){

        try{
            JMenuItem item;

            item  = new JMenuItem("�d��");
            item.setLocale(new Locale("ja","JP"));
            item.setFont(new java.awt.Font("dialog", 0, 18));
            item.addActionListener(new CMSPower());
            obj.add(item);

            obj.addSeparator();
            item  = new JMenuItem("���䃂�[�h");
            item.setLocale(new Locale("ja","JP"));
            item.setFont(new java.awt.Font("dialog", 0, 18));
            item.addActionListener(new CMSControlMode());
            obj.add(item);

            obj.addSeparator();
            item  = new JMenuItem("�v���Z�X");
            item.setLocale(new Locale("ja","JP"));
            item.setFont(new java.awt.Font("dialog", 0, 18));
            item.addActionListener(new CMSProcChg());
            obj.add(item);

            obj.addSeparator();
            item  = new JMenuItem("�V�[�h���x");
            item.setLocale(new Locale("ja","JP"));
            item.setFont(new java.awt.Font("dialog", 0, 18));
            item.addActionListener(new CMSSeedSpeed());
            obj.add(item);

            item  = new JMenuItem("�V�[�h��]");
            item.setLocale(new Locale("ja","JP"));
            item.setFont(new java.awt.Font("dialog", 0, 18));
            item.addActionListener(new CMSSeedRotation());
            obj.add(item);

            item  = new JMenuItem("�V�[�h�ʒu");
            item.setLocale(new Locale("ja","JP"));
            item.setFont(new java.awt.Font("dialog", 0, 18));
            item.addActionListener(new CMSSeedPosition());
            obj.add(item);

            obj.addSeparator();
            item  = new JMenuItem("���c�{���x");
            item.setLocale(new Locale("ja","JP"));
            item.setFont(new java.awt.Font("dialog", 0, 18));
            item.addActionListener(new CMSCrucibleSpeed());
            obj.add(item);

            item  = new JMenuItem("���c�{��]");
            item.setLocale(new Locale("ja","JP"));
            item.setFont(new java.awt.Font("dialog", 0, 18));
            item.addActionListener(new CMSCrucibleRotation());
            obj.add(item);

            item  = new JMenuItem("���c�{�ʒu");
            item.setLocale(new Locale("ja","JP"));
            item.setFont(new java.awt.Font("dialog", 0, 18));
            item.addActionListener(new CMSCruciblePosition());
            obj.add(item);

            obj.addSeparator();
            item  = new JMenuItem("�����ێ�");
            item.setLocale(new Locale("ja","JP"));
            item.setFont(new java.awt.Font("dialog", 0, 18));
            item.addActionListener(new CMSSXLHoldPosition());
            obj.add(item);

            obj.addSeparator();
            item  = new JMenuItem("���C���q�[�^�P�d��");
            item.setLocale(new Locale("ja","JP"));
            item.setFont(new java.awt.Font("dialog", 0, 18));
            item.addActionListener(new CMSHeater1Power());
            obj.add(item);

            item  = new JMenuItem("���C���q�[�^�Q�d��");
            item.setLocale(new Locale("ja","JP"));
            item.setFont(new java.awt.Font("dialog", 0, 18));
            item.addActionListener(new CMSHeater2Power());
            obj.add(item);

            item  = new JMenuItem("�{�g���q�[�^�d��");
            item.setLocale(new Locale("ja","JP"));
            item.setFont(new java.awt.Font("dialog", 0, 18));
            item.addActionListener(new CMSHeaterBottomPower());
            obj.add(item);

            item  = new JMenuItem("�q�[�^���x");
            item.setLocale(new Locale("ja","JP"));
            item.setFont(new java.awt.Font("dialog", 0, 18));
            item.addActionListener(new CMSHeaterTemp());
            obj.add(item);

            item  = new JMenuItem("�V�[�h�q�[�^�d��");
            item.setLocale(new Locale("ja","JP"));
            item.setFont(new java.awt.Font("dialog", 0, 18));
            item.addActionListener(new CMSHeaterSeedPower());
            obj.add(item);

            obj.addSeparator();
            item  = new JMenuItem("�v���A���S��");
            item.setLocale(new Locale("ja","JP"));
            item.setFont(new java.awt.Font("dialog", 0, 18));
            item.addActionListener(new CMSArPull());
            obj.add(item);

            item  = new JMenuItem("�g�b�v�A���S��");
            item.setLocale(new Locale("ja","JP"));
            item.setFont(new java.awt.Font("dialog", 0, 18));
            item.addActionListener(new CMSArTop());
            obj.add(item);
/*@@ ���ꋭ�x
            obj.addSeparator();
            item  = new JMenuItem("���ꋭ�x�P");
            item.setLocale(new Locale("ja","JP"));
            item.setFont(new java.awt.Font("dialog", 0, 18));
            obj.add(item);

            item  = new JMenuItem("���ꋭ�x�Q");
            item.setLocale(new Locale("ja","JP"));
            item.setFont(new java.awt.Font("dialog", 0, 18));
            obj.add(item);

            obj.addSeparator();
            item  = new JMenuItem("����v���Z�X");
            item.setLocale(new Locale("ja","JP"));
            item.setFont(new java.awt.Font("dialog", 0, 18));
            obj.add(item);
@@*/
            obj.addSeparator();
            item  = new JMenuItem("���j�^�[�ؑւ�");
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

    // ----- ��������ActionListener ------------------------

    /*******************************************************
     *
     *   �e�[�u��
     *
     *******************************************************/
    //
    // ----- ���ƒ萔�ݒ� ----------------------------------
    //
    class OperationTable implements ActionListener {

        private CZOperationTable obj = null;

        public void actionPerformed(ActionEvent e){
            CZSystem.getOperationMst();     //@@@@ ���ƒ萔�}�X�^��Ǎ���
            if(null == obj) obj = new CZOperationTable();
            	if( 0 != CZSystemDefine.TIMER_FLG ){
            		obj.timerStart();
            	}
            obj.setDefault();
            obj.setVisible(true);
        }
    }

    // ----- ����e�[�u���ݒ� ------------------------------
    //
    class ControlTable implements ActionListener {

        private CZControlTable obj = null;

        public void actionPerformed(ActionEvent e){
          CZSystem.getControlMst();     //@@@@ ����e�[�u���}�X�^��Ǎ���
          if(null == obj) obj = new CZControlTable();
          	if( 0 != CZSystemDefine.TIMER_FLG ){
	      		obj.timerStart();
	        }
          obj.setDefault();
          obj.setVisible(true);
        }
    }

    // ----- ���ƃe�[�u����r --2006.06.02 Y.K-------------
    //
    class TblHikaku1 implements ActionListener {

        private CZTblHikaku obj = null;

        public void actionPerformed(ActionEvent e){
          if(null == obj) obj = new CZTblHikaku(1);
//		  obj.setType(1);		//1:���Ɓ@2:����
//          obj.setDefault();
          obj.setVisible(true);
        }
    }

    // ----- ����e�[�u����r --2006.06.02 Y.K-------------
    //
    class TblHikaku2 implements ActionListener {

        private CZTblHikaku2 obj = null;

        public void actionPerformed(ActionEvent e){
          if(null == obj) obj = new CZTblHikaku2();
          obj.setVisible(true);
        }
    }

    // ----- ���V�s���e�o�� ---------------
    //
    class TblRcpOutPut implements ActionListener {

        private CZTblRcpOutPut obj = null;

        public void actionPerformed(ActionEvent e){
          if(null == obj) obj = new CZTblRcpOutPut();
          obj.setVisible(true);
        }
    }
// add start 2008.09.29
    // ----- �ύX�����o�� ---------------
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

    // ----- �I�� ------------------------------------------
    //
    class Exit implements ActionListener {
        public void actionPerformed(ActionEvent e){
            CZSystem.exit(0,"Menu Exit");
        }
    }

    /*******************************************************
     *
     *   ����
     *
     *******************************************************/
    //
    // ----- �o�u�ۑ� --------------------------------------
    //
    class PVData implements ActionListener {
        private CZPVDataSave obj = null;

        public void actionPerformed(ActionEvent e){
            if(null == obj) obj = new CZPVDataSave();
            obj.setDefault();
            obj.setVisible(true);

        }
    }


    // ----- �s�o�f ----------------------------------------
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

    // ----- �g�����h�e�[�u�� ------------------------------
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

    // ----- �e���`���� ------------------------------------
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
    // ----- �����o�u�ۑ� --------------------------------------
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
     *   �v��
     *
     *******************************************************/
    //
    // ----- �b�b�c���g�` ----------------------------------
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

    // ----- �b�b�c�摜 ------------------------------------
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
    // ----- �P�x�ω��`�F�b�N ------------------------------
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
     *   �G���[
     *
     *******************************************************/
    //
    // ----- �G���[�\�� ------------------------------------
    //
    class ErrorMsgWin implements ActionListener {
        private CZErrorMsgWin obj = null;

        public void actionPerformed(ActionEvent e){
            if(null == obj) obj = new CZErrorMsgWin();
            obj.setDefault();
            obj.setVisible(true);
        }
    }

    // ----- �T�[�o�[�G���[�\�� ----------------------------
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
     *   �����e�i���X
     *
     *******************************************************/
    //
    // ----- �G���[���ڐݒ� --------------------------------
    //
    class ErrorSetWin implements ActionListener {
        private CZErrorSetWin obj = null;

        public void actionPerformed(ActionEvent e){
            if(null == obj) obj = new CZErrorSetWin();
            obj.setDefault();
            obj.setVisible(true);
        }
    }

    // ----- ���ƒ萔�R�s�[ --------------------------------
    //
    class OperationTableCp implements ActionListener {
        private CZOperationTableCp obj = null;

        public void actionPerformed(ActionEvent e){
            if(null == obj) obj = new CZOperationTableCp();
            obj.setDefault();
            obj.setVisible(true);
        }
    }

    // ----- �l�n�Ǘ� --------------------------------------
    //
    class MOControl implements ActionListener {
        private CZMOControl obj = null;

        public void actionPerformed(ActionEvent e){
            if(null == obj) obj = new CZMOControl();
            obj.setDefault();
            obj.setVisible(true);
        }
    }

    // ----- �q�`�h�c��� ----------------------------------
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
     *   ���
     *
     *******************************************************/
    //
    // ----- �ғ��� --------------------------------------
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
    // ----- �Ď��� --------------------------------------
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
    // ----- �r���� --------------------------------------
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
     *   �W���Ď�
     *
     *******************************************************/
    //
    // ----- �d�� ------------------------------------------
    //
    class CMSPower implements ActionListener {
        private CZCMSPower obj = null;

        public void actionPerformed(ActionEvent e){
            if(null == obj) obj = new CZCMSPower();
            obj.setDefault();
            obj.setVisible(true);
        }
    }

    // ----- ���䃂�[�h ------------------------------------
    //
    class CMSControlMode implements ActionListener {
        private CZCMSControlMode obj = null;

        public void actionPerformed(ActionEvent e){
            if(null == obj) obj = new CZCMSControlMode();
            obj.setDefault();
            obj.setVisible(true);
        }
    }

    // ----- �v���Z�X�ڍs ----------------------------------
    //
    class CMSProcChg implements ActionListener {
        private CZCMSProcChg obj = null;

        public void actionPerformed(ActionEvent e){
            if(null == obj) obj = new CZCMSProcChg();
            obj.setDefault();
            obj.setVisible(true);
        }
    }

    // ----- �V�[�h���x ------------------------------------
    //
    class CMSSeedSpeed implements ActionListener {
        private CZCMSSeedSpeed obj = null;

        public void actionPerformed(ActionEvent e){
            if(null == obj) obj = new CZCMSSeedSpeed();
            obj.setDefault();
            obj.setVisible(true);
        }
    }

    // ----- �V�[�h��] ------------------------------------
    //
    class CMSSeedRotation implements ActionListener {
        private CZCMSSeedRotation obj = null;

        public void actionPerformed(ActionEvent e){
            if(null == obj) obj = new CZCMSSeedRotation();
            obj.setDefault();
            obj.setVisible(true);
        }
    }

    // ----- �V�[�h�ʒu ------------------------------------
    //
    class CMSSeedPosition implements ActionListener {
        private CZCMSSeedPosition obj = null;

        public void actionPerformed(ActionEvent e){
            if(null == obj) obj = new CZCMSSeedPosition();
            obj.setDefault();
            obj.setVisible(true);
        }
    }

    // ----- ���c�{���x ------------------------------------
    //
    class CMSCrucibleSpeed implements ActionListener {
        private CZCMSCrucibleSpeed obj = null;

        public void actionPerformed(ActionEvent e){
            if(null == obj) obj = new CZCMSCrucibleSpeed();
            obj.setDefault();
            obj.setVisible(true);
        }
    }

    // ----- ���c�{��] ------------------------------------
    //
    class CMSCrucibleRotation implements ActionListener {
        private CZCMSCrucibleRotation obj = null;

        public void actionPerformed(ActionEvent e){
            if(null == obj) obj = new CZCMSCrucibleRotation();
            obj.setDefault();
            obj.setVisible(true);
        }
    }

    // ----- ���c�{�ʒu ------------------------------------
    //
    class CMSCruciblePosition implements ActionListener {
        private CZCMSCruciblePosition obj = null;

        public void actionPerformed(ActionEvent e){
            if(null == obj) obj = new CZCMSCruciblePosition();
            obj.setDefault();
            obj.setVisible(true);
        }
    }

    // ----- �����ێ��ʒu ----------------------------------
    //
    class CMSSXLHoldPosition implements ActionListener {
        private CZCMSSXLHoldPosition obj = null;

        public void actionPerformed(ActionEvent e){
            if(null == obj) obj = new CZCMSSXLHoldPosition();
            obj.setDefault();
            obj.setVisible(true);
        }
    }

    // ----- �q�[�^�[�P�p���[ ------------------------------
    //
    class CMSHeater1Power implements ActionListener {
        private CZCMSHeater1Power obj = null;

        public void actionPerformed(ActionEvent e){
            if(null == obj) obj = new CZCMSHeater1Power();
            obj.setDefault();
            obj.setVisible(true);
        }
    }

    // ----- �q�[�^�[�Q�p���[ ------------------------------
    //
    class CMSHeater2Power implements ActionListener {
        private CZCMSHeater2Power obj = null;

        public void actionPerformed(ActionEvent e){
            if(null == obj) obj = new CZCMSHeater2Power();
            obj.setDefault();
            obj.setVisible(true);
        }
    }

    // ----- �{�g���q�[�^�[�p���[ --------------------------
    //
    class CMSHeaterBottomPower implements ActionListener {
        private CZCMSHeaterBottomPower obj = null;

        public void actionPerformed(ActionEvent e){
            if(null == obj) obj = new CZCMSHeaterBottomPower();
            obj.setDefault();
            obj.setVisible(true);
        }
    }

    // ----- �q�[�^�[�P���x --------------------------------
    //
    class CMSHeaterTemp implements ActionListener {
        private CZCMSHeaterTemp obj = null;

        public void actionPerformed(ActionEvent e){
            if(null == obj) obj = new CZCMSHeaterTemp();
            obj.setDefault();
            obj.setVisible(true);
        }
    }

    // ----- �V�[�h�q�[�^�[�d�� ----------------------------
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
    // ----- �v���A���S�� ----------------------------------
    //
    class CMSArPull implements ActionListener {
        private CZCMSArPull obj = null;

        public void actionPerformed(ActionEvent e){
            if(null == obj) obj = new CZCMSArPull();
            obj.setDefault();
            obj.setVisible(true);
        }
    }

    // ----- �g�b�v�A���S�� --------------------------------
    //
    class CMSArTop implements ActionListener {
        private CZCMSArTop obj = null;

        public void actionPerformed(ActionEvent e){
            if(null == obj) obj = new CZCMSArTop();
            obj.setDefault();
            obj.setVisible(true);
        }
    }

    // ----- ���j�^�[�ؑւ� --------------------------------
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
