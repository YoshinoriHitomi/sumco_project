package cz;

import java.util.Locale;

import javax.swing.JTable;
import javax.swing.ListSelectionModel;
import javax.swing.table.DefaultTableColumnModel;
import javax.swing.table.TableColumn;

import czclass.CZNativeHikiage;

/*******************************************************************************
 *
 *   引き上げ条件用スクロールパネル内テーブル
 * @author  (KPK Co.,Ltd.)
 * @version 1.0 (2003/04/01)
 *  T6追加に伴う修正 @@
 *******************************************************************************/
public class CZSetPanelSetTbl extends JTable {

    private CZSetPanelSetTblMdl model = null;

    CZSetPanelSetTbl(){
        super();

        try{
            setName("CZSetPanelSetTbl");
            setBounds(0, 0, 200, 200);
            setAutoCreateColumnsFromModel(true);
            setSelectionMode(ListSelectionModel.SINGLE_SELECTION);
            setLocale(new Locale("ja","JP"));
            setFont(new java.awt.Font("dialog", 0, 12));
            setRowHeight(17);

            model = new CZSetPanelSetTblMdl();

            setModel(model);
            DefaultTableColumnModel cmdl = (DefaultTableColumnModel)getColumnModel();
            TableColumn  colum = null;

            colum = cmdl.getColumn(0);
            colum.setMaxWidth(60);
            colum.setMinWidth(60);
            colum.setWidth(60); 
            colum.setCellRenderer(new CZSetPanelSetTblRenderer());

            colum = cmdl.getColumn(1);
            colum.setMaxWidth(190);
            colum.setMinWidth(190);
            colum.setWidth(190);
            colum.setCellRenderer(new CZSetPanelSetTblRenderer());

//@@            CZSystem.log("CZSetPanelSetTbl","new");
        }
        catch (Throwable e) {
            CZSystem.handleException(e);
        }
    }

    //
    // 引き上げ条件変更
    //
    public void alterTbl(){

        CZNativeHikiage tbl = CZSystem.getBtSet();
//@@        CZSystem.log("CZSetPanelSetTbl alterTbl","TBL["+tbl.getBatch()+"]");

        int i = 0;  // バッチNo 
        model.setValueAt(tbl.getBatch(),i,1);

        i++;        // PG-ID
        model.setValueAt(tbl.getPgid(),i,1);

        i++;        // 品種
        model.setValueAt(tbl.getHinshu(),i,1);

        i++;        // 方位
        model.setValueAt(tbl.getHoui(),i,1);

        i++;        // タイプ
        model.setValueAt(tbl.getH_type(),i,1);

        i++;        // 比抵抗
        model.setValueAt(tbl.getHiteikou(),i,1);

        i++;        // 酸素
        model.setValueAt(tbl.getSanso(),i,1);

        i++;        // GAP
        model.setValueAt(tbl.getGap(),i,1);

        i++;        // ルツボ径
        model.setValueAt(new Integer(tbl.getRutubo_kei()),i,1);

        i++;        // プルアルゴン
        model.setValueAt(new Integer(tbl.getPull_ar()),i,1);

        i++;        // トップアルゴン
        model.setValueAt(new Integer(tbl.getTop_ar()),i,1);

        i++;        // 直径
        model.setValueAt(new Integer(tbl.getChokkei()),i,1);

        i++;        // 引上長
        model.setValueAt(new Integer(tbl.getHikiage_cho()),i,1);

        i++;        // 初期仕込量
        model.setValueAt(new Integer(tbl.getI_sikomi()),i,1);

        i++;        // 追加仕込量
        model.setValueAt(new Integer(tbl.getT_sikomi()),i,1);

        i++;        // 残液量
        model.setValueAt(new Integer(tbl.getZaneki()),i,1);

        i++;        // レシピーＮｏ（溶解）
        model.setValueAt(new Integer(tbl.getNo_youkai()),i,1);

        i++;        // レシピーＮｏ（引上）
        model.setValueAt(new Integer(tbl.getNo_hikiage()),i,1);

        i++;        // レシピーＮｏ（回転）
        model.setValueAt(new Integer(tbl.getNo_kaiten()),i,1);

        i++;        // レシピーＮｏ（取出）
        model.setValueAt(new Integer(tbl.getNo_toridasi()),i,1);

        i++;        // レシピーＮｏ（圧力）
        model.setValueAt(new Integer(tbl.getNo_aturyoku()),i,1);

        i++;        // レシピーＮｏ（定数
        model.setValueAt(new Integer(tbl.getNo_teisu()),i,1);

        repaint();
    }
}           
