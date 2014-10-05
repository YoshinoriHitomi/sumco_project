package cz;

import java.io.Serializable;

/**
 *  ÇbÇbÇcê∂îgå`ÉfÅ[É^
 * @author  (KPK Co.,Ltd.)
 * @version 1.0 (2003/04/01)
 */
public class CZSystemCCDWave implements Serializable 
{
    String  s_time;
    String  batch;

    int p_no;
    int sp_no;
    int p_renban;
    int p_time;
    int sp_time;

    String  slice;

    int s_start;
    int s_end;

    float   single;
    float   k_chokei;
    float   h_chokei;
    int v_keisoku;
    int h_keisoku;

    String  status;
    String  route;
    String  cross;
    String  search;
    String  peek;
    String  hosei;

    int len;

    String  data;
}
