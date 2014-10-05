package cz;

import java.io.Serializable;

/**
 *  操業ＰＶ実績管理
 * @author  (KPK Co.,Ltd.)
 * @version 1.0 (2003/10/21)
 */
public class CZSystemPvControl implements Serializable 
{
    public String   batch;          //バッチ番号
    public String	t_name;			//テーブル名
    public String	s_start;		//採取開始日時
    public String	s_end;			//採取終了日時
    public int		m_flg;			//間引き有無
    public int		m_sumi;			//間引き済
    public int		mo_flg;			//ＭＯ保存フラグ
    public String	mo_date;		//ＭＯ保存日時
}
