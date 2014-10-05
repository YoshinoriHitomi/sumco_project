package cz;

import java.io.Serializable;

//==========================================================================
/**
*  移動平均計算結果データ（１レコード分）
*/
public class CalcData implements Serializable
{
    public  float   fp_ave;
    public  float   pf_ave;
    public  float   pf_umax_ave;
    public  float   pf_max_ave;
    public  float   pf_lmin_ave;
    public  float   pf_min_ave;
    public  int     judg;
} //CalcData
