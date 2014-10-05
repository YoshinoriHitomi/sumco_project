package cz;

import java.io.Serializable;

/**
 *  エラーメッセージ
 * @author  (KPK Co.,Ltd.)
 * @version 1.0 (2003/04/01)
 */
public class CZSystemErr implements Serializable 
{
    public  int e_no;
    public  String  o_time;
    public  String  batch;
    public  int p_no;
    public  int sp_no;
    public  int p_renban;
    public  int p_time;
    public  int sp_time;
    public  int flg_error;
    public  int info1;
    public  int info2;
    public  String  ro_info;
    public  String  ban_info;
    public  String  k_time;
}