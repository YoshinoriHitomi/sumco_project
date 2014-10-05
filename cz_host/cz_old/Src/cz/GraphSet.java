package cz;

import java.io.Serializable;
import java.awt.Color;

//==========================================================================
/**
 *  グラフ表示用条件データ
 */
public class GraphSet implements Serializable
{
    public  Color   fp_umax_col;
    public  Color   fp_max_col;
    public  Color   fp_min_col;
    public  Color   fp_lmin_col;
    public  Color   fp_umax_over_col;
    public  Color   fp_max_over_col;
    public  Color   fp_center_col;
    public  Color   fp_min_over_col;
    public  Color   fp_lmin_over_col;

    public  boolean shld_shift;
    public  float   shld_shift_val;

    public  boolean fp_pf_ave_draw;
    public  boolean fp_draw;
    public  boolean fp_pf_draw;
    public  boolean dia_draw;
    public  boolean dia_pf_draw;
    public  boolean sxl_rpm_draw;
    public  boolean cru_rpm_draw;

    public  Color   fp_pf_ave_draw_col;
    public  Color   fp_draw_col;
    public  Color   fp_pf_draw_col;
    public  Color   dia_draw_col;
    public  Color   dia_pf_draw_col;
    public  Color   sxl_rpm_draw_col;
    public  Color   cru_rpm_draw_col;

    public  float   x_min;
    public  float   x_max;

    public  int     x_bun;
    public  int     x_koushi;
    public  int     x_memkan;
    public  int     x_memketa;
    public  int     x_syouketa;

    public  float   y_min;
    public  float   y_max;

    public  int     y_bun;
    public  int     y_koushi;
    public  int     y_memkan;
    public  int     y_memketa;
    public  int     y_syouketa;

    public  float   y_dia_min;
    public  float   y_dia_max;

    public  float   y_rpm_min;
    public  float   y_rpm_max;
} //GraphSet
