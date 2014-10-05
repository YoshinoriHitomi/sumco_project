package cz;

import java.io.Serializable;

/**
 *  制御テーブル項目定義
 * @author  (KPK Co.,Ltd.)
 * @version 1.0 (2003/04/01)
 */
public class CZSystemCtName implements Serializable 
{
    public  int g_no;
    public  int t_no;
    public  String  t_name;
    public  String  l_name;
    public  String  l_unit;
    public  float   l_min;
    public  float   l_max;
    public  int l_keta;
    public  String  r_name;
    public  String  r_unit;
    public  float   r_min;
    public  float   r_max;
    public  int r_keta;
    public  int k_sort;
    public  int pv_no;

}
