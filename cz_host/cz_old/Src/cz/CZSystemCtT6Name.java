package cz;

import java.io.Serializable;

/**
 *  制御テーブル項目定義(T6)
 * @author  (KPK Co.,Ltd.)
 * @version 1.0 (2003/04/01)
 */
public class CZSystemCtT6Name implements Serializable 
{
    public  int     g_no;
    public  int     k_no1;
    public  int     k_no2;
    public  int     k_no;
    public  String  k_name;
    public  String  k_unit;
    public  float   k_min;
    public  float   k_max;
    public  int     k_keta;
    public  int     k_sort;
    public  int     pv_no;

}
