package cz;

import java.io.Serializable;

/**
 *  エラーメッセージ
 * @author  (KPK Co.,Ltd.)
 * @version 1.0 (2003/04/01)
 */
public class CZSystemHostErr implements Serializable 
{
    public  int e_no;
    public  String  o_time;
    public  int p_no;
    public  int info1;
    public  int info2;
    public  String  mname;
    public  String  k_time;
}
