package cz;

import javax.swing.table.AbstractTableModel;

/**
 *   ＰＶデータ表示用スクロールパネル内テーブルのモデル  
 * @author  (KPK Co.,Ltd.)
 * @version 1.0 (2003/04/01)
 */      

public class CZSetPanelPVTblMdl extends AbstractTableModel {

    final int TBL_COL   = 5;
    final int TBL_ROW   = CZPV.PV_DATA_SET_LENGTH;

    final Object data[][] = new Object[TBL_ROW][TBL_COL];

    final String[] names = {"#","Ch","Name","Data","Unit"};

    CZSetPanelPVTblMdl(){
        super(); 

        try{
            int i = 0;
            data[i][0] = new Integer(i+1);
            data[i][1] = new String("1");
            data[i][2] = new String("###.##");
            data[i][3] = new String("9999.9999");
            data[i][4] = new String("###/###");

            i++; 
            data[i][0] = new Integer(i+1);
            data[i][1] = new String("2");
            data[i][2] = new String("###.##");
            data[i][3] = new String("9999.9999");
            data[i][4] = new String("###/###");

            i++; 
            data[i][0] = new Integer(i+1);
            data[i][1] = new String("3");
            data[i][2] = new String("###.##");
            data[i][3] = new String("9999.9999");
            data[i][4] = new String("###/###");

            i++; 
            data[i][0] = new Integer(i+1);
            data[i][1] = new String("4");
            data[i][2] = new String("###.##");
            data[i][3] = new String("9999.9999");
            data[i][4] = new String("###/###");

            i++; 
            data[i][0] = new Integer(i+1);
            data[i][1] = new String("5");
            data[i][2] = new String("###.##");
            data[i][3] = new String("9999.9999");
            data[i][4] = new String("###/###");

            i++; 
            data[i][0] = new Integer(i+1);
            data[i][1] = new String("6");
            data[i][2] = new String("###.##");
            data[i][3] = new String("9999.9999");
            data[i][4] = new String("###/###");

            i++; 
            data[i][0] = new Integer(i+1);
            data[i][1] = new String("7");
            data[i][2] = new String("###.##");
            data[i][3] = new String("9999.9999");
            data[i][4] = new String("###/###");

            i++; 
            data[i][0] = new Integer(i+1);
            data[i][1] = new String("8");
            data[i][2] = new String("###.##");
            data[i][3] = new String("9999.9999");
            data[i][4] = new String("###/###");

            i++; 
            data[i][0] = new Integer(i+1);
            data[i][1] = new String("9");
            data[i][2] = new String("###.##");
            data[i][3] = new String("9999.9999");
            data[i][4] = new String("###/###");

            i++; 
            data[i][0] = new Integer(i+1);
            data[i][1] = new String("10");
            data[i][2] = new String("###.##");
            data[i][3] = new String("9999.5");
            data[i][4] = new String("###/###");

        }
        catch (Throwable e) {
            CZSystem.handleException(e); 
        }
//@@        CZSystem.log("CZSetPanelPVTblMdl","new"); 
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
