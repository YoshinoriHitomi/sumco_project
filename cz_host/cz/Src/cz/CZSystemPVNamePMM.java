package cz;

import java.io.Serializable;

/**
 *  ‰^“]‰æ–Ê’è‹`
 * @author  (KPK Co.,Ltd.)
 * @version 1.0 (2003/04/01)
 */
public class CZSystemPVNamePMM implements Serializable 
{
    static final    int SIZE        = 10;

    static final    int X_TIME      = 1;
    static final    int X_LENGTH    = 2;

    int p_no;

    int     item[]  = new int[SIZE];
    float   min[]   = new float[SIZE];
    float   max[]   = new float[SIZE];

    int x_shubetu; // 1:ŽžŠÔ 2:’·‚³

    int x_time;
    int x_width;

}
