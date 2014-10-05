package cz;

import java.io.Serializable;

/**
 *  テーブル変更履歴
 * @author  (KPK Co.,Ltd.)
 * @version 1.0 (2003/04/01)
 */
public class CZSystemTblModify implements Serializable
{
    public  String  s_time;
    public  String  op_name;
    public  String  batch;
    public  String  message;

    public  int     key1;
    public  int     key2;
    public  int     key3;
}
