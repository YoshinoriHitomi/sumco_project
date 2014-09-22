package cz;

import java.io.Serializable;

/**
 *  輝度変化チェックデータ
 * @author  (KPK Co.,Ltd.)
 * @version 1.0 (2003/04/01)
 */
public class CZSystemBrightnessData implements Serializable 
{
    String  s_time;
    String  batch;
    int     p_no;
    int     charge;
    String  gap;
    float   max_b_ave;
    float   range_b_ave;
    float   max_b_judge;
    float   range_b_judge;
    float   x_review;
    float   review_range;
    float   body_l_max_b_ave;
    float   body_r_max_b_ave;
    float   body_max_b_range;
    float   body_peek;
    float   body_peek_judge;
    int     len;
    String  data;
    String  c_batch;
    float   c_max_b_ave;
    float   c_range_b_ave;
    float   t_max_b_judge;
    float   t_range_b_judge;
    float   c_body_l_max_b_ave;
    float   c_body_r_max_b_ave;
    float   t_body_l_max_b_ave;
    float   t_body_r_max_b_ave;
}
