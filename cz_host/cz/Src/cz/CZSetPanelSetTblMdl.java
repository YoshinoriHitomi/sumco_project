package cz;

import javax.swing.table.AbstractTableModel;

/*******************************************************************************
 *
 *   �����グ�����p�X�N���[���p�l�����e�[�u���̃��f��
 * @author  (KPK Co.,Ltd.)
 * @version 1.0 (2003/04/01)
 * T6�ǉ��ɔ����C�� @@
 *******************************************************************************/
public class CZSetPanelSetTblMdl extends AbstractTableModel { 

    final int TBL_COL   = 2;
    final int TBL_ROW   = 22;       //@@ 21 -> 22

    final Object data[][] = new Object[TBL_ROW][TBL_COL];

    final String[] names = {"����", "���e"};

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
            data[i][0] = new String("�i��");
            data[i][1] = new String("XX-XXXX");

            i++;
            data[i][0] = new String("����");
            data[i][1] = new String("XXX");

            i++;
            data[i][0] = new String("�^�C�v");
            data[i][1] = new String("XX");

            i++;
            data[i][0] = new String("���R");
            data[i][1] = new String("XXXXX.XXXXX-XXXXX.XXXXX");

            i++;
            data[i][0] = new String("�_�f");
            data[i][1] = new String("XX.XX-XX.XX");

            i++;
            data[i][0] = new String("GAP");
            data[i][1] = new String("321");

            i++;
            data[i][0] = new String("���c�{");
            data[i][1] = new String("XX");

            i++;
            data[i][0] = new String("�v��Ar");
            data[i][1] = new String("XXX.XX");

            i++;
            data[i][0] = new String("�g�b�vAr");
            data[i][1] = new String("XXX.XX");

            i++;
            data[i][0] = new String("���a");
            data[i][1] = new String("XXX.XX");

            i++;
            data[i][0] = new String("���㒷");
            data[i][1] = new String("XXXX.X");

            i++;
            data[i][0] = new String("�d����");
            data[i][1] = new String("XXXXXX");

            i++;
            data[i][0] = new String("�d����(R)");
            data[i][1] = new String("XXXXXX");

            i++;
            data[i][0] = new String("�c�t");
            data[i][1] = new String("XXXXXX");

            i++;
            data[i][0] = new String("�n��(T1)");
            data[i][1] = new String("XXX");

            i++;
            data[i][0] = new String("����(T2)");
            data[i][1] = new String("XXX");

            i++;
            data[i][0] = new String("��](T3)");
            data[i][1] = new String("XXX");

            i++;
            data[i][0] = new String("��o(T4)");
            data[i][1] = new String("XXX");

            i++;
            data[i][0] = new String("����(T5)");
            data[i][1] = new String("XXX");

            i++;
            data[i][0] = new String("�萔(T6)");    //@@
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
