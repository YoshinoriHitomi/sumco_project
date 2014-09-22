package cz;

import java.awt.Color;

/***********************************************************
 *   �o�u�֌W 
 *       ���сA�O���t�ݒ�
 * @author  (KPK Co.,Ltd.)
 * @version 1.0 (2003/04/01)
 * @Update 2013/10/30  �\���؂�ւ��@�\ (@20131030)
 ***********************************************************/
public class CZPV {

    // ----- �O���t�̉����ݒ� ------------------------------
    public static final int TIME_SCALE_ST   = 5;
    public static final int LENGTH_SCALE_ST = 5;

    private static boolean  time_flag       = false;
    private static int      time_scale      = 120;
    private static int      length_scale    = 20;

    // ----- �O���t�̍ő�f�[�^�� --------------------------
    public static final int PV_MAX_COUNT  = 65535;

    // ----- �O���t�\���G���A�̕����� ----------------------
    public static final int PV_X_SPLIT = 6;
    public static final int PV_Y_SPLIT = 5;

    // ----- �O���t�̐F ------------------------------------
    public static final Color PV_BACK_COLOR    = java.awt.Color.black;
    public static final Color PV_MEM_COLOR     = java.awt.Color.white;
    public static final Color PV_MEM_SP1_COLOR = java.awt.Color.gray;
    public static final Color PV_MEM_SP2_COLOR = java.awt.Color.darkGray;

    public static final Color PV_COLOR[] ={
                                    java.awt.Color.orange,
                                    java.awt.Color.red,
                                    java.awt.Color.yellow,
                                    java.awt.Color.cyan,
                                    java.awt.Color.green,
                                    java.awt.Color.white,
                                    java.awt.Color.white,
                                    java.awt.Color.white,
                                    java.awt.Color.white,
                                    java.awt.Color.white };

    // @20131030
    public static final Color PV_COLOR2[] ={
                                    java.awt.Color.white,
                                    java.awt.Color.white,
                                    java.awt.Color.white,
                                    java.awt.Color.white,
                                    java.awt.Color.white,
                                    java.awt.Color.orange,
                                    java.awt.Color.red,
                                    java.awt.Color.yellow,
                                    java.awt.Color.cyan,
                                    java.awt.Color.green };

    // ----- �����グ�����̐ݒ� ----------------------------
    public  static final int PV_DATA_SET_LENGTH    = 10; 
    public  static final int PV_DATA_SET_GR_LENGTH = 10;	// @20131030
    private static int    pv_DATA_SET[]     = new int[PV_DATA_SET_LENGTH];
    private static int    pv_DATA_SET_CH[]  = new int[PV_DATA_SET_LENGTH];
    private static String pv_DATA_SET_NAME[]= new String[PV_DATA_SET_LENGTH];
    private static float  pv_DATA_SET_DATA[]= new float[PV_DATA_SET_LENGTH];
    private static String pv_DATA_SET_UNIT[]= new String[PV_DATA_SET_LENGTH];
    private static float  pv_DATA_SET_MM[][]= new float[PV_DATA_SET_LENGTH][2];

    // ----- PV�f�[�^�̃f�[�^�� ----------------------------
    public  static final int PV_MAX_LENGTH = 128;
//@@@33    public  static final int PV_MAX_LENGTH = 10;

    // ----- PV�f�[�^����M�� ----------------------------
    private static float pv_DATA[] = new float[PV_MAX_LENGTH];

    // ----- PV�f�[�^ DB����̓ǂݍ��ݕ� -------------------
    private static float pv_DATA_DB[][];

    // ----- PV�f�[�^ �V�K������ ---------------------------
    private static float pv_DATA_USE[][];

    // ----- PV SXL���� ------------------------------------
    private static final int SXL_L    = 4;

    private static int   pv_count_db  = -1;
    private static int   pv_count_use = -1;

    // ----- �q�[�^�[ON���� --------------------------------
    private static int   pv_HT_ON_TIME;

    // *****************************************************
    // ----- ���������� ------------------------------------
    // *****************************************************
    public static synchronized boolean newCZPV(){
        try{ 
            newPVDataDB();
            newPVDataUse();
            return true; 
        }
        catch (Throwable e) {
            CZSystem.handleException(e); 
            return false;
        }
    }

    // *****************************************************
    // ----- PV�f�[�^ DB���̋L���̈揉���� -----------------
    // *****************************************************
    private static synchronized int newPVDataDB(){ 
        pv_DATA_DB = new float[PV_MAX_COUNT][12];	 // @20131030
        pv_count_db  = 0;
        return pv_count_db;
    }

    // *****************************************************
    // ----- PV�f�[�^ DB���ǉ� -----------------------------
    // *****************************************************
    public static int addPVDataDB(
                    float tim,float len,float p1,float p2,float p3,float p4,float p5,float p6,float p7,float p8,float p9,float p10){
        pv_DATA_DB[pv_count_db][0] = tim;
        pv_DATA_DB[pv_count_db][1] = len;
        pv_DATA_DB[pv_count_db][2] = p1; 
        pv_DATA_DB[pv_count_db][3] = p2; 
        pv_DATA_DB[pv_count_db][4] = p3; 
        pv_DATA_DB[pv_count_db][5] = p4; 
        pv_DATA_DB[pv_count_db][6] = p5; 
        pv_DATA_DB[pv_count_db][7] = p6; 	// @20131030
        pv_DATA_DB[pv_count_db][8] = p7; 	// @20131030
        pv_DATA_DB[pv_count_db][9] = p8; 	// @20131030
        pv_DATA_DB[pv_count_db][10] = p9; 	// @20131030
        pv_DATA_DB[pv_count_db][11] = p10; 	// @20131030
        pv_count_db++;
        return pv_count_db;
    }
     
    // *****************************************************
    // ----- PV�f�[�^ DB���̃f�[�^ -------------------------
    // *****************************************************
    public static float getPVDataDB(int no,int pos){ 
        return pv_DATA_DB[no][pos];
    }
     
    // *****************************************************
    // PV�f�[�^ DB���̃f�[�^�� -----------------------------
    // *****************************************************
    public static int getPVDataDBCount(){
        return pv_count_db;
    }

    // *****************************************************
    // ----- PV�f�[�^ �V�K�������̋L���̈揉���� -----------
    // *****************************************************
    private static synchronized int newPVDataUse(){
        pv_DATA_USE = new float[PV_MAX_COUNT][12];	// @20131030
        pv_count_use = 0;
        return pv_count_use;
    }

    // *****************************************************
    // ----- PV�f�[�^ �V�K�������ǉ� -----------------------
    // *****************************************************
    public static synchronized int addPVDataUse(
            String bt,  int p_no,int sp_no,  int p_renban,int p_time, int sp_time,
            int p_date, int h_ontime,int hk_renban, float data[]){

        if(0 > pv_count_use) return pv_count_use;

        pv_HT_ON_TIME = h_ontime;

        for(int i = 0 ; i < PV_MAX_LENGTH ; i++){
            pv_DATA[i] = data[i];
        }
        for(int i = 0 ; i < PV_DATA_SET_LENGTH ; i++){ 
            pv_DATA_SET_DATA[i] = data[pv_DATA_SET_CH[i] - 1];
        }
        pv_DATA_USE[pv_count_use][0] = (float)p_time;
        pv_DATA_USE[pv_count_use][1] = data[SXL_L];
        for(int i = 0 ; i < PV_DATA_SET_GR_LENGTH ; i++){
            pv_DATA_USE[pv_count_use][i+2] = pv_DATA_SET_DATA[i];
        }

//@@        CZSystem.log("CZPV addPVDataUse","[" + pv_count_use + "][" + 
//@@                    pv_DATA_USE[pv_count_use][0] + "][" +
//@@                    pv_DATA_USE[pv_count_use][1] + "][" +
//@@                    pv_DATA_USE[pv_count_use][2] + "][" +
//@@                    pv_DATA_USE[pv_count_use][3] + "][" +
//@@                    pv_DATA_USE[pv_count_use][4] + "][" +
//@@                    pv_DATA_USE[pv_count_use][5] + "][" +
//@@                    pv_DATA_USE[pv_count_use][6] + "][" +
//@@                    pv_DATA_USE[pv_count_use][7] + "][" +
//@@                    pv_DATA_USE[pv_count_use][8] + "][" +
//@@                    pv_DATA_USE[pv_count_use][9] + "][" +
//@@                    pv_DATA_USE[pv_count_use][10] + "]" );

//@@        CZSystem.log("CZPV addPVDataUse","[" + pv_count_use + "][" + 
//@@                    pv_DATA_SET_DATA[0] + "][" + 
//@@                    pv_DATA_SET_DATA[1] + "][" + 
//@@                    pv_DATA_SET_DATA[2] + "][" + 
//@@                    pv_DATA_SET_DATA[3] + "][" + 
//@@                    pv_DATA_SET_DATA[4] + "][" + 
//@@                    pv_DATA_SET_DATA[5] + "][" + 
//@@                    pv_DATA_SET_DATA[6] + "][" + 
//@@                    pv_DATA_SET_DATA[7] + "][" + 
//@@                    pv_DATA_SET_DATA[8] + "][" + 
//@@                    pv_DATA_SET_DATA[9] + "]" ); 
        pv_count_use++;
        return pv_count_use; 
    }

    // *****************************************************
    // ----- PV�f�[�^ �V�K�������̃f�[�^ -------------------
    // *****************************************************
    public static float getPVDataUse(int no,int pos){
        return pv_DATA_USE[no][pos]; 
    }

    // *****************************************************
    // ----- PV�f�[�^ �V�K�������̃f�[�^�� -----------------
    // *****************************************************
    public static int getPVDataUseCount(){ 
        return pv_count_use; 
    }

    // *****************************************************
    // ----- PV�f�[�^�e�[�u���\���p�i���݂�PV�f�[�^) -------
    // *****************************************************
    public static float getPVDataSet(int i){ 
        return pv_DATA_SET_DATA[i];
    }

    // *****************************************************
    // ----- PV�f�[�^�i���݂�PV�f�[�^) ---------------------
    // *****************************************************
    public static float getPVData(int i){
        return pv_DATA[i];
    }

    // *****************************************************
    // ----- �q�[�^�[ON���� --------------------------------
    // *****************************************************
    public static int getHtOnTm(){ 
        return pv_HT_ON_TIME;
    }

    // *****************************************************
    // �O���t�o�u���� --------------------------------------
    // *****************************************************
    public static int getPVGrNo(int no){ 
        return pv_DATA_SET_CH[no];
    }
    // *****************************************************
    // ----- �O���t�o�u���ڂ̐ݒ� --------------------------
    // *****************************************************
    public static boolean setPVGrNo(int no[]){ 
        for(int i = 0 ; i < PV_DATA_SET_LENGTH ; i++){ 
            pv_DATA_SET_CH[i] = no[i];
//@@            CZSystem.log("CZPV setPVGrNo","[" + i + "][" + pv_DATA_SET_CH[i] + "]"); 
        }
        return true; 
    }

    // *****************************************************
    // ----- �O���t�������Ԃ����� --------------------------
    // *****************************************************
    public static boolean getPVGrTimeFlag(){ 
        return time_flag;
    }

    // *****************************************************
    // ----- �O���t�������Ԃ������̐ݒ� --------------------
    // *****************************************************
    public static boolean setPVGrTimeFlag(boolean flag){ 
        time_flag = flag;
        return time_flag;
    }

    // *****************************************************
    // ----- �O���t�������Ԃ̔{�� --------------------------
    // *****************************************************
    public static int getPVGrTimeScale(){
        return time_scale;
    }

    // *****************************************************
    // ----- �O���t�������Ԃ̔{���ݒ� ----------------------
    // *****************************************************
    public static int setPVGrTimeScale(int val){ 
        if(1 > val) return time_scale;
        time_scale = val;
        return time_scale;
    }

    // *****************************************************
    // ----- �O���t�������Ԃ̔{�� --------------------------
    // *****************************************************
    public static int getPVGrLengthScale(){
        return length_scale; 
    }

    // *****************************************************
    // ----- �O���t�������Ԃ̔{���ݒ� ----------------------
    // *****************************************************
    public static int setPVGrLengthScale(int val){ 
        if(1 > val) return length_scale; 
        length_scale = val;
        return length_scale; 
    }

    // *****************************************************
    // ----- �O���t�c���̂����� ----------------------------
    // *****************************************************
    public static float getPVGrMin(int no){
        float ret = (float)0;
        try{ 
            ret = pv_DATA_SET_MM[no][0]; 
        }    
        catch (Throwable e) {
            CZSystem.handleException(e);
            CZSystem.exit(-1,"getPVGrMin Error[" + no + "]");
        }
        return ret;
    }

    // *****************************************************
    // ----- �O���t�c���̂������̐ݒ� ----------------------
    // *****************************************************
    public static boolean setPVGrMin(float val[]){
        try{ 
            for(int i = 0 ; i < val.length ; i++)
                pv_DATA_SET_MM[i][0] = val[i];
        }
        catch (Throwable e) {
            CZSystem.handleException(e);
            CZSystem.exit(-1,"setPVGrMin Error[" + val + "]");
        }
        return true; 
    }

    // *****************************************************
    // ----- �O���t�c���̂����� ----------------------------
    // *****************************************************
    public static float getPVGrMax(int no){
        float ret = (float)0;
        try{ 
            ret = pv_DATA_SET_MM[no][1]; 
        }    
        catch (Throwable e) {
            CZSystem.handleException(e);
            CZSystem.exit(-1,"getPVGrMax Error[" + no + "]");
        }
        return ret;
    }

    // *****************************************************
    // ----- �O���t�c���̂������̐ݒ� ----------------------
    // *****************************************************
    public static boolean setPVGrMax(float val[]){ 
        try{ 
            for(int i = 0 ; i < val.length ; i++)
            pv_DATA_SET_MM[i][1] = val[i];  
        }    
        catch (Throwable e) {
            CZSystem.handleException(e);
            CZSystem.exit(-1,"setPVGrMax Error[" + val + "]");
        }
        return true; 
    }

    // *****************************************************
    // ----- PV�f�[�^�̖��O --------------------------------
    // *****************************************************
    public static String getPVGrName(int no){
        try{ 
            return pv_DATA_SET_NAME[no]; 
        }    
            catch (Throwable e) {
            CZSystem.handleException(e);
            CZSystem.exit(-1,"getPVGrName Error[" + no + "]");
        }
        return null; 
    }

    // *****************************************************
    // ----- PV�f�[�^�̖��O�ݒ� ----------------------------
    // *****************************************************
    public static boolean setPVGrName(String val[]){ 
        try{ 
            for(int i = 0 ; i < val.length ; i++)
            pv_DATA_SET_NAME[i]  = val[i];
        }    
        catch (Throwable e) {
            CZSystem.handleException(e);
            CZSystem.exit(-1,"setPVGrName Error[" + val + "]");
        }
        return true; 
    }

    // *****************************************************
    // ----- PV�f�[�^�̒P�� --------------------------------
    // *****************************************************
    public static String getPVGrUnit(int no){
        try{ 
            return pv_DATA_SET_UNIT[no];
        }    
        catch (Throwable e) {
            CZSystem.handleException(e);
            CZSystem.exit(-1,"getPVGrUnit Error[" + no + "]");
        }
        return null; 
    }

    // *****************************************************
    // ----- PV�f�[�^�̒P�ʂ̐ݒ� --------------------------
    // *****************************************************
    public static boolean setPVGrUnit(String val[]){ 
        try{ 
            for(int i = 0 ; i < val.length ; i++)
            pv_DATA_SET_UNIT[i]  = val[i];
        }    
        catch (Throwable e) {
            CZSystem.handleException(e);
            CZSystem.exit(-1,"setPVGrUnit Error[" + val + "]");
        }
        return true; 
    }
}    
