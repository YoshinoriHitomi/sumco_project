package cz;

import java.io.Serializable;

/**
 *  引き上げ条件
 * @author  (KPK Co.,Ltd.)
 * @version 1.0 (2003/04/01)
 */
public class CZSystemBtTemp implements Serializable 
{
    public String   batch;          //バッチ番号
    public String   pgid;           //PG-ID
    public String   t_time;         //登録日時
    public int      renban;         //連番
    public String   hinshu;         //品種
    public String   houi;           //方位
    public String   h_type;         //タイプ
    public String   hiteikou;       //比抵抗
    public String   sanso;          //酸素
    public String   gap;            //GAP
    public int      rutubo_kei;     //ルツボ径
    public int      chokkei;        //直径
    public int      hikiage_cho;    //引上長
    public int      top_ar;         //トップアルゴン
    public int      pull_ar;        //プルアルゴン
    public int      i_sikomi;       //仕込量
    public int      t_sikomi;       //追加仕込量
    public int      zaneki;         //残液量
    public int      no_youkai;      //T1(溶解)
    public int      no_hikiage;     //T2(引上)
    public int      no_kaiten;      //T3(回転)
    public int      no_toridasi;    //T4(取出)
    public int      no_aturyoku;    //T5(圧力)
    public int      no_teisu;       //T6(定数) @@
    public int      pno_start;      //スタートプロセス
    public int      p_kaisi;        //開始
}
