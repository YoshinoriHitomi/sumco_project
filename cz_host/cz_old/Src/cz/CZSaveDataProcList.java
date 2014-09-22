package cz;

import java.io.Serializable;

/**
 *  複数PVデータ保存用プロセスNoリスト
 * @author  (KPK Co.,Ltd.)
 * @version 1.0 (2003/04/01)
 */
public class CZSaveDataProcList implements Serializable 
{
    public int      p_no;            //プロセスNo
    public int      sp_no;           //サブプロセスNo
    public int      p_renban;        //プロセス連番
}
