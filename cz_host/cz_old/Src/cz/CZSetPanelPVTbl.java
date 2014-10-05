package cz;

import java.util.Locale;

import javax.swing.JTable;
import javax.swing.ListSelectionModel;
import javax.swing.table.DefaultTableColumnModel;
import javax.swing.table.TableColumn;

/**
 *   ＰＶ表示項目用スクロールパネル内テーブル
 * @author  (KPK Co.,Ltd.)
 * @version 1.0 (2003/04/01)
 */  
public class CZSetPanelPVTbl extends JTable {
    
    private CZSetPanelPVTblMdl model = null;
    
    CZSetPanelPVTbl(){
        super();
            
        try{
            setName("CZSetPanelPVTbl");
            setBounds(0, 0, 200, 200);
            setAutoCreateColumnsFromModel(true);
            setSelectionMode(ListSelectionModel.SINGLE_SELECTION);
            setLocale(new Locale("ja","JP"));
            setFont(new java.awt.Font("dialog", 0, 12));
            setRowHeight(17);

            model = new CZSetPanelPVTblMdl();
            setModel(model);
            DefaultTableColumnModel cmdl = (DefaultTableColumnModel)getColumnModel();
            TableColumn  colum = null;

            colum = cmdl.getColumn(0);
            colum.setMaxWidth(20);
            colum.setMinWidth(20);
            colum.setWidth(20);
            colum.setCellRenderer(new CZSetPanelPVTblRenderer());

            colum = cmdl.getColumn(1);
            colum.setMaxWidth(30);
            colum.setMinWidth(30);
            colum.setWidth(30);
            colum.setCellRenderer(new CZSetPanelPVTblRenderer());

            colum = cmdl.getColumn(2);
            colum.setMaxWidth(70);
            colum.setMinWidth(70);
            colum.setWidth(70);
            colum.setCellRenderer(new CZSetPanelPVTblRenderer());

            colum = cmdl.getColumn(3);
            colum.setMaxWidth(75);
            colum.setMinWidth(75);
            colum.setWidth(75);
            colum.setCellRenderer(new CZSetPanelPVTblRenderer());

            colum = cmdl.getColumn(4);
            colum.setMaxWidth(55);
            colum.setMinWidth(55);
            colum.setWidth(55);
            colum.setCellRenderer(new CZSetPanelPVTblRenderer());
//@@            CZSystem.log("CZSetPanelPVTbl","new");
        }
        catch (Throwable e) {
            CZSystem.handleException(e);
        }
    }
            
    //
    // PVデータ数値更新 テーブルパネル１０個分
    //
    public void alterPV(){

//@@        CZSystem.log("CZSetPanelPVTbl alterPV","1");
        for(int i = 0 ; i < 10 ; i++){
            model.setValueAt(new Integer(CZPV.getPVGrNo(i)),i,1);
            model.setValueAt(new String(CZPV.getPVGrName(i).trim()),i,2);
            model.setValueAt(new Float(CZPV.getPVDataSet(i)),i,3);
            model.setValueAt(new String(CZPV.getPVGrUnit(i).trim()),i,4);
        }

        repaint();
    }
}
