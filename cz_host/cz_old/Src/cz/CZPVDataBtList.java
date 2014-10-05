package cz;

import java.io.Serializable;

/**
 *  PV引き上げバッチリスト
 * @author  (KPK Co.,Ltd.)
 * @version 1.0 (2003/04/01)
 */
public class CZPVDataBtList implements Serializable 
{
    public int      flg;            //Checkフラグ
    public String   batch;          //バッチ番号
    public String   hinshu;         //品種
    public int      i_sikomi;       //仕込量
    public int      no_hikiage;     //T2(引上)
}
