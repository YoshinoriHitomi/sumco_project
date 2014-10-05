package cz;

import java.awt.Component;
import java.util.Locale;

import javax.swing.JTable;
import javax.swing.table.DefaultTableCellRenderer;

/**
 *   ＰＶデータ用スクロールパネル内テーブルのレンダラー
 * @author  (KPK Co.,Ltd.)
 * @version 1.0 (2003/04/01)
 */  
class CZSetPanelPVTblRenderer extends DefaultTableCellRenderer {
    public Component getTableCellRendererComponent( JTable table,
                                                Object value,
                                                boolean isSelected,
                                                boolean hasFocus,
                                                int row,int column){

        boolean untenflg = true;

        setLocale(new Locale("ja","JP"));
        setFont(new java.awt.Font("dialog", 0, 12));

        setValue(value);

        setBackground(CZPV.PV_BACK_COLOR);

        untenflg = CZSystem.untenView();
        if (untenflg == true){
	        if(CZPV.PV_COLOR.length <= row) setForeground(java.awt.Color.cyan);
	        else setForeground(CZPV.PV_COLOR[row]);
        } else {
	        if(CZPV.PV_COLOR2.length <= row) setForeground(java.awt.Color.cyan);
	        else setForeground(CZPV.PV_COLOR2[row]);
        }

        return(this);
    }
}
