package cz;

import java.util.Locale;

import javax.swing.JTable;
import javax.swing.ListSelectionModel;
import javax.swing.table.DefaultTableColumnModel;
import javax.swing.table.TableColumn;

import czclass.CZNativeHikiage;

/*******************************************************************************
 *
 *   �����グ�����p�X�N���[���p�l�����e�[�u��
 * @author  (KPK Co.,Ltd.)
 * @version 1.0 (2003/04/01)
 *  T6�ǉ��ɔ����C�� @@
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
    // �����グ�����ύX
    //
    public void alterTbl(){

        CZNativeHikiage tbl = CZSystem.getBtSet();
//@@        CZSystem.log("CZSetPanelSetTbl alterTbl","TBL["+tbl.getBatch()+"]");

        int i = 0;  // �o�b�`No 
        model.setValueAt(tbl.getBatch(),i,1);

        i++;        // PG-ID
        model.setValueAt(tbl.getPgid(),i,1);

        i++;        // �i��
        model.setValueAt(tbl.getHinshu(),i,1);

        i++;        // ����
        model.setValueAt(tbl.getHoui(),i,1);

        i++;        // �^�C�v
        model.setValueAt(tbl.getH_type(),i,1);

        i++;        // ���R
        model.setValueAt(tbl.getHiteikou(),i,1);

        i++;        // �_�f
        model.setValueAt(tbl.getSanso(),i,1);

        i++;        // GAP
        model.setValueAt(tbl.getGap(),i,1);

        i++;        // ���c�{�a
        model.setValueAt(new Integer(tbl.getRutubo_kei()),i,1);

        i++;        // �v���A���S��
        model.setValueAt(new Integer(tbl.getPull_ar()),i,1);

        i++;        // �g�b�v�A���S��
        model.setValueAt(new Integer(tbl.getTop_ar()),i,1);

        i++;        // ���a
        model.setValueAt(new Integer(tbl.getChokkei()),i,1);

        i++;        // ���㒷
        model.setValueAt(new Integer(tbl.getHikiage_cho()),i,1);

        i++;        // �����d����
        model.setValueAt(new Integer(tbl.getI_sikomi()),i,1);

        i++;        // �ǉ��d����
        model.setValueAt(new Integer(tbl.getT_sikomi()),i,1);

        i++;        // �c�t��
        model.setValueAt(new Integer(tbl.getZaneki()),i,1);

        i++;        // ���V�s�[�m���i�n���j
        model.setValueAt(new Integer(tbl.getNo_youkai()),i,1);

        i++;        // ���V�s�[�m���i����j
        model.setValueAt(new Integer(tbl.getNo_hikiage()),i,1);

        i++;        // ���V�s�[�m���i��]�j
        model.setValueAt(new Integer(tbl.getNo_kaiten()),i,1);

        i++;        // ���V�s�[�m���i��o�j
        model.setValueAt(new Integer(tbl.getNo_toridasi()),i,1);

        i++;        // ���V�s�[�m���i���́j
        model.setValueAt(new Integer(tbl.getNo_aturyoku()),i,1);

        i++;        // ���V�s�[�m���i�萔
        model.setValueAt(new Integer(tbl.getNo_teisu()),i,1);

        repaint();
    }
}           
