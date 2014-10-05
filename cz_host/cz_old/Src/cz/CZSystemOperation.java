package cz;

import java.io.Serializable;

/**
 *      オペレータ介入操作
 * @author  (KPK Co.,Ltd.)
 * @version 1.0 (2003/04/01)
 */
public class CZSystemOperation implements Serializable
{
    public  String  s_time;
    public  String  batch;
    public  String  p_name;

    public  int     p_renban;
    public  int     p_time;

    public  String  message;

    public  int     sid;

    public  int     val1;
    public  int     val2;
    public  int     val3;

}
