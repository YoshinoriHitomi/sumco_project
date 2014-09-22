package cz;

import java.io.Serializable;

/**
 *  運転画面定義
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

    int x_shubetu; // 1:時間 2:長さ

    int x_time;
    int x_width;

}
