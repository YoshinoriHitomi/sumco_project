package cz;

import java.io.Serializable;

/**
 *  引き上げ開始
 * @author  (KPK Co.,Ltd.)
 * @version 1.0 (2003/04/01)
 */
public class CZSystemStart implements Serializable 
{
    public String   batch;
    public int      p_no;
    public int      sp_no;
    public int      p_renban;
    public String   p_start;
}
