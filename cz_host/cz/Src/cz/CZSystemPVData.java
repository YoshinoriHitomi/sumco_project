package cz;

import java.io.Serializable;

/**
 *  ‚o‚uƒf[ƒ^
 * @author  (KPK Co.,Ltd.)
 * @version 1.0 (2003/04/01)
 */
public class CZSystemPVData implements Serializable 
{
    public int      p_no;
    public int      sp_no;
    public int      p_renban;
    public int      p_time;
    public int      sp_time;
    public String   p_date;
    public int      h_ontime;
    public int      hk_renban;

    public float    p_length;

    public float    data[]  = new float[CZSystemDefine.PV_MAX_LENGTH];
}
