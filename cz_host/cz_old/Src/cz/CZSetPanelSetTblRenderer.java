package cz;

import java.awt.Component;
import java.util.Locale;

import javax.swing.JTable;
import javax.swing.table.DefaultTableCellRenderer;

/*******************************************************************************
 *
 *   引き上げ条件用スクロールパネル内テーブルのレンダラー
 * @author  (KPK Co.,Ltd.)
 * @version 1.0 (2003/04/01)
 *******************************************************************************/  
class CZSetPanelSetTblRenderer extends DefaultTableCellRenderer {
    public Component getTableCellRendererComponent( JTable table,
                                                Object value,
                                                boolean isSelected,
                                                boolean hasFocus,
                                                int row,int column){
        setLocale(new Locale("ja","JP"));
        setFont(new java.awt.Font("dialog", 0, 12));
        setValue(value);
        if(1 == column){
            setForeground(java.awt.Color.black);
            return(this);
        }
        if(16 > row) setForeground(java.awt.Color.red);
        else setForeground(java.awt.Color.blue);
        return(this);
     }
}
