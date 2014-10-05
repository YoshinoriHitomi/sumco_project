package cz;

import javax.swing.table.AbstractTableModel;

/*******************************************************************************
 *
 *   引き上げ条件用スクロールパネル内テーブルのモデル
 * @author  (KPK Co.,Ltd.)
 * @version 1.0 (2003/04/01)
 * T6追加に伴う修正 @@
 *******************************************************************************/
public class CZSetPanelSetTblMdl extends AbstractTableModel { 

    final int TBL_COL   = 2;
    final int TBL_ROW   = 22;       //@@ 21 -> 22

    final Object data[][] = new Object[TBL_ROW][TBL_COL];

    final String[] names = {"項目", "内容"};

    CZSetPanelSetTblMdl(){ 
        super();

        try{
            int i = 0;
            data[i][0] = new String("BtNo");
            data[i][1] = new String("XXXC-XXXB");

            i++;
            data[i][0] = new String("PG-ID");
            data[i][1] = new String("12345678");

            i++;
            data[i][0] = new String("品種");
            data[i][1] = new String("XX-XXXX");

            i++;
            data[i][0] = new String("方位");
            data[i][1] = new String("XXX");

            i++;
            data[i][0] = new String("タイプ");
            data[i][1] = new String("XX");

            i++;
            data[i][0] = new String("比抵抗");
            data[i][1] = new String("XXXXX.XXXXX-XXXXX.XXXXX");

            i++;
            data[i][0] = new String("酸素");
            data[i][1] = new String("XX.XX-XX.XX");

            i++;
            data[i][0] = new String("GAP");
            data[i][1] = new String("321");

            i++;
            data[i][0] = new String("ルツボ");
            data[i][1] = new String("XX");

            i++;
            data[i][0] = new String("プルAr");
            data[i][1] = new String("XXX.XX");

            i++;
            data[i][0] = new String("トップAr");
            data[i][1] = new String("XXX.XX");

            i++;
            data[i][0] = new String("直径");
            data[i][1] = new String("XXX.XX");

            i++;
            data[i][0] = new String("引上長");
            data[i][1] = new String("XXXX.X");

            i++;
            data[i][0] = new String("仕込み");
            data[i][1] = new String("XXXXXX");

            i++;
            data[i][0] = new String("仕込み(R)");
            data[i][1] = new String("XXXXXX");

            i++;
            data[i][0] = new String("残液");
            data[i][1] = new String("XXXXXX");

            i++;
            data[i][0] = new String("溶解(T1)");
            data[i][1] = new String("XXX");

            i++;
            data[i][0] = new String("引上(T2)");
            data[i][1] = new String("XXX");

            i++;
            data[i][0] = new String("回転(T3)");
            data[i][1] = new String("XXX");

            i++;
            data[i][0] = new String("取出(T4)");
            data[i][1] = new String("XXX");

            i++;
            data[i][0] = new String("圧力(T5)");
            data[i][1] = new String("XXX");

            i++;
            data[i][0] = new String("定数(T6)");    //@@
            data[i][1] = new String("XXX");

        }
        catch (Throwable e) {
            CZSystem.handleException(e);
        }
//@@        CZSystem.log("CZSetPanelSetTblMdl","new");
    }

    public int getColumnCount(){
        return TBL_COL;
    }

    public int getRowCount(){
        return TBL_ROW;
    }

    public Object getValueAt(int row, int col){
        return data[row][col];
    }

    public String getColumnName(int column){
        return names[column];
    }

    public Class getColumnClass(int c){
        return getValueAt(0, c).getClass();
    }

    public boolean isCellEditable(int row, int col){
        return false;
    }

    public void setValueAt(Object aValue, int row, int column){ 
        data[row][column] = aValue;
    }
}
