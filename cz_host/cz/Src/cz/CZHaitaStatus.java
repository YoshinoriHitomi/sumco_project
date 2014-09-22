package cz;

import java.awt.Color;
import java.awt.Component;
import java.awt.Rectangle;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.text.DecimalFormat;
import java.util.Locale;

import java.util.Properties;

import java.io.FileInputStream;

import javax.swing.JButton;
import javax.swing.JDialog;
import javax.swing.JLabel;
import javax.swing.JScrollPane;
import javax.swing.JTable;
import javax.swing.JPanel;
import javax.swing.ListSelectionModel;
import javax.swing.plaf.metal.MetalBorders.Flush3DBorder;
import javax.swing.table.AbstractTableModel;
import javax.swing.table.DefaultTableCellRenderer;
import javax.swing.table.DefaultTableColumnModel;
import javax.swing.table.TableColumn;

import czclass.CZNativeRoState;
import czclass.CZNativeCTState;
import czclass.CZNativeSTState;
import czclass.CZNativeRoHikiage;

/**
 *   �ғ��󋵕\�����
 * @author  (KPK Co.,Ltd.)
 * @version 1.0 (2003/04/01)
 */
public class CZHaitaStatus extends JDialog {

    private final int   RO_MAX = 99;
//    private final int   RO_MAX = 50;

    private JLabel      time_lab        = null; //�X�V�����\��

    private JScrollPane st_scpanel      = null; //�X�N���[��
    private StatusTable st_table        = null; //�r���󋵃e�[�u��

    private CZNativeRoState[] status_list;
    private CZNativeCTState[] ctstatus_list;
    private CZNativeSTState[] ststatus_list;
    private CZNativeRoHikiage[] rohikiage_list;

	private titlPanel ct_panel;
	private titlPanel st_panel;

    private UpdateThread    updateTh    = null; //�X�V�X���b�h

    private final int   IP_LIST_CNT = 50;

    private String prop_IP[];
    private String prop_Memo[];

    /**
    *   �R���X�g���N�^
    */
    CZHaitaStatus(){
        super();

        setTitle("�r����");
        setSize(1150,700);
        setResizable(false);
        setModal(false);

        getContentPane().setLayout(null);
        // ����n�Q�Ƌ@�\    @20131021
        if(CZSystemDefine.REFERENCE_RUN == CZSystem.getRunLevel()){
            getContentPane().setBackground(CZSystemDefine.DEFAULT_REFERENCE_BACKGROUND_COL);
        }else{
            getContentPane().setBackground(CZSystemDefine.DEFAULT_BACKGROUND_COL);
        }

        JLabel lab = new JLabel("�X�V����",JLabel.CENTER);
        lab.setBounds(20, 20, 100, 24);
        lab.setLocale(new Locale("ja","JP"));
        lab.setFont(new java.awt.Font("dialog", 0, 18));
        lab.setBorder(new Flush3DBorder());
        lab.setForeground(java.awt.Color.black);
        getContentPane().add(lab);

        time_lab = new JLabel("",JLabel.CENTER);
        time_lab.setBounds(120, 20, 200, 24);
        time_lab.setLocale(new Locale("ja","JP"));
        time_lab.setFont(new java.awt.Font("dialog", 0, 18));
        time_lab.setBorder(new Flush3DBorder());
        time_lab.setForeground(java.awt.Color.black);
        getContentPane().add(time_lab);

		titlPanel ct_panel = new titlPanel("����e�[�u��", new Rectangle( 0, 0, 275,20));
		ct_panel.setBounds(240, 50, 275, 20);
		getContentPane().add(ct_panel);

		titlPanel st_panel = new titlPanel("���ƒ萔", new Rectangle( 0, 0, 275, 20));
		st_panel.setBounds(515, 50, 275, 20);
		getContentPane().add(st_panel);

        //
        st_table = new StatusTable();
        st_scpanel = new JScrollPane(st_table);
        st_scpanel.setBounds(20, 70, 1110, 565);
        st_scpanel.setBorder(new Flush3DBorder());
        getContentPane().add(st_scpanel);

        updateTh = new UpdateThread();
        updateTh.setPriority(Thread.MIN_PRIORITY);
        updateTh.start();
        
        CZSystem.log("CZHaitaStatus","�r���󋵁@�\��");

        try{
            // ----- Property_File��� IP�A�h���X�E�N���C�A���gPC�ݒu�����擾����B --------
            Properties prop =  new Properties();
            FileInputStream pros = new FileInputStream("IP_PROPERTY(EUC).TXT");
            prop.load(pros);

            // IP�̐ݒ�
            prop_IP  = new String[IP_LIST_CNT];
            prop_Memo = new String[IP_LIST_CNT];
            for(int i=0; i < IP_LIST_CNT ; i++){
                try {
                    prop_IP[i]   = prop.getProperty("C" + (i+1) + "_IP_NO");
                    prop_Memo[i]  = prop.getProperty("C" + (i+1) + "_MEMO");
                } catch (Exception e) {
                    prop_IP[i]   = new String("");
                    prop_Memo[i]  = new String("");
                }
            }
        } catch( Exception e ) {
                                        //�v���p�e�B�擾�ŃG���[�̎��́A�I������B
CZSystem.log("CZHaitaStatus","NO Propertie File");
            // CZSystem.exit(-1,"CZHaitaStatus NO Propertie File");
        }

    }


    //
    //
    //
    public boolean setDefault(){

        return true;
    }

    //
    //
    //
    private boolean updateStatus(){

        status_list = null;
        ctstatus_list = null;
        ststatus_list = null;
        rohikiage_list = null;

        status_list = CZSystem.CZNativeRoStateGet();
        ctstatus_list = CZSystem.CZNativeCTStateGet();
        ststatus_list = CZSystem.CZNativeSTStateGet();
        rohikiage_list = CZSystem.CZNativeRoHikiageGet();

        if(null == status_list) return false;

        int size = status_list.length;
        if(1 > size) return false;

        st_table.setData();

        String tm = CZSystem.getDateTime();
        time_lab.setText(tm);

        return true;
    } 


    /**
    *   �X�V�X���b�h
    */
    class UpdateThread extends Thread {

        UpdateThread(){

        }


        public void run(){
//@@            CZSystem.log("CZHaitaStatus","UpdateThread START");

            while(true){
                updateStatus();
                CZSystem.sleep(10000);
            } // while end
        }
    }


    /**
    *   �󋵕\���p�e�[�u��
    */
    public class StatusTable extends JTable {

        private StatusModel model = null;
            
        StatusTable(){
            super();

            setName("StatusTable");
            setAutoCreateColumnsFromModel(true);
            setSelectionMode(ListSelectionModel.SINGLE_SELECTION);
            setLocale(new Locale("ja","JP"));
            setFont(new java.awt.Font("dialog", 0, 12));
            setRowHeight(17);

            model = new StatusModel();
            setModel(model);

            DefaultTableColumnModel cmdl = (DefaultTableColumnModel)getColumnModel();
            TableColumn     colum = null;
            ColorRender ren   = null;

            //No
            colum = cmdl.getColumn(0);
            colum.setMaxWidth(25);
            colum.setMinWidth(25);
            colum.setWidth(25);

            //�F
            ren = new ColorRender();
            ren.setHorizontalAlignment(ren.CENTER);
            colum = cmdl.getColumn(1);
            colum.setMaxWidth(40);
            colum.setMinWidth(40);
            colum.setWidth(40);
            colum.setCellRenderer(ren);

            //BtNo
            ren = new ColorRender();
            colum = cmdl.getColumn(2);
            colum.setMaxWidth(90);
            colum.setMinWidth(90);
            colum.setWidth(90);
            colum.setCellRenderer(ren);

            //�v���Z�X
            ren = new ColorRender();
            colum = cmdl.getColumn(3);
            colum.setMaxWidth(65);
            colum.setMinWidth(65);
            colum.setWidth(65);
            colum.setCellRenderer(ren);

            //Status�i�r���j�i����e�[�u���p�j
            ren = new ColorRender();
            ren.setHorizontalAlignment(ren.CENTER);
            colum = cmdl.getColumn(4);
            colum.setMaxWidth(40);
            colum.setMinWidth(40);
            colum.setWidth(40);
            colum.setCellRenderer(ren);

            //IP_Address�i����e�[�u���p�j
            ren = new ColorRender();
            ren.setHorizontalAlignment(ren.CENTER);
            colum = cmdl.getColumn(5);
            colum.setMaxWidth(110);
            colum.setMinWidth(110);
            colum.setWidth(110);
            colum.setCellRenderer(ren);

            //���l�i����e�[�u���p�j
            ren = new ColorRender();
            ren.setHorizontalAlignment(ren.CENTER);
            colum = cmdl.getColumn(6);
            colum.setMaxWidth(125);
            colum.setMinWidth(125);
            colum.setWidth(125);
            colum.setCellRenderer(ren);

            //Status�i�r���j�i���ƒ萔�p�j
            ren = new ColorRender();
            ren.setHorizontalAlignment(ren.CENTER);
            colum = cmdl.getColumn(7);
            colum.setMaxWidth(40);
            colum.setMinWidth(40);
            colum.setWidth(40);
            colum.setCellRenderer(ren);

            //IP_Address�i���ƒ萔�p�j
            ren = new ColorRender();
            ren.setHorizontalAlignment(ren.CENTER);
            colum = cmdl.getColumn(8);
            colum.setMaxWidth(110);
            colum.setMinWidth(110);
            colum.setWidth(110);
            colum.setCellRenderer(ren);

            //���l�i���ƒ萔�p�j
            ren = new ColorRender();
            ren.setHorizontalAlignment(ren.CENTER);
            colum = cmdl.getColumn(9);
            colum.setMaxWidth(125);
            colum.setMinWidth(125);
            colum.setWidth(125);
            colum.setCellRenderer(ren);

            //�v���Z�X����
            ren = new ColorRender();
            ren.setHorizontalAlignment(ren.CENTER);
            colum = cmdl.getColumn(10);
            colum.setMaxWidth(80);
            colum.setMinWidth(80);
            colum.setWidth(80);
            colum.setCellRenderer(ren);

            //�q�[�^�[�I������
            ren = new ColorRender();
            ren.setHorizontalAlignment(ren.CENTER);
            colum = cmdl.getColumn(11);
            colum.setMaxWidth(80);
            colum.setMinWidth(80);
            colum.setWidth(80);
            colum.setCellRenderer(ren);

            //�i��
            ren = new ColorRender();
            ren.setHorizontalAlignment(ren.CENTER);
            colum = cmdl.getColumn(12);
            colum.setMaxWidth(80);
            colum.setMinWidth(80);
            colum.setWidth(80);
            colum.setCellRenderer(ren);

            //PG-ID
            ren = new ColorRender();
            ren.setHorizontalAlignment(ren.CENTER);
            colum = cmdl.getColumn(13);
            colum.setMaxWidth(80);
            colum.setMinWidth(80);
            colum.setWidth(80);
            colum.setCellRenderer(ren);

        }

        //
        //
        //
        public void setData(){

            if(null == status_list) return;
            int size = status_list.length;
            if(1 > size) return;

            DecimalFormat f1 = new DecimalFormat("0");

            ColorRender cell;

            for(int i = 0 ; i < size ; i++){
                CZNativeRoState st = status_list[i];
                CZNativeCTState cts = ctstatus_list[i];
                CZNativeSTState sts = ststatus_list[i];
                CZNativeRoHikiage rh = rohikiage_list[i];

                //�F��
				String s = CZSystem.RoKetaChg(st.getRoName());	// 20050725 �F�F�\�������ύX
                model.setValueAt(s, i,1);
//                model.setValueAt(st.getRoName(), i,1);

                //�o�b�`No
                model.setValueAt(st.getBatch(), i,2);

                //�v���Z�X
                model.setValueAt(new ColorString(CZSystem.getProcName(st.getP_no()),java.awt.Color.blue), i,3);
                if(st.getDown()) model.setValueAt(new ColorString("DOWN",java.awt.Color.red), i,3);

                //Status�i����e�[�u���p�j
                model.setValueAt(f1.format(cts.getExclusive()), i,4);

                //IP-Address�i����e�[�u���p�j
                model.setValueAt(cts.getAdds(), i,5);

                //���l�i����e�[�u���p�j
                for(int p = 0; p < IP_LIST_CNT ; p++){
                    if( cts.getAdds().equals(prop_IP[p])){
                        model.setValueAt(prop_Memo[p], i,6);
                        break;
                    }else{
                        model.setValueAt("", i,6);
                    }
                }


                //Status�i���ƒ萔�p�j
                model.setValueAt(f1.format(sts.getExclusive()), i,7);

                //IP-Address�i���ƒ萔�p�j
                model.setValueAt(sts.getAdds(), i,8);

                //���l�i���ƒ萔�p�j
                for(int p = 0; p < IP_LIST_CNT ; p++){
                    if( sts.getAdds().equals(prop_IP[p])){
                        model.setValueAt(prop_Memo[p], i,9);
                        break;
                    }else{
                        model.setValueAt("", i,9);
                    }
                }
                

                //�v���Z�X����
                model.setValueAt(CZSystem.timeFormat(st.getP_time()), i,10);

                //�q�[�^�I������
                if(0 >= st.getH_ontime()) {
                    model.setValueAt(
                        new ColorString(CZSystem.timeFormat(st.getH_ontime()),java.awt.Color.red), i,11);
                } else {
                    model.setValueAt(
                        new ColorString(CZSystem.timeFormat(st.getH_ontime()),java.awt.Color.blue), i,11);
                }

                //�i��
                model.setValueAt(rh.getHinshu(), i,12);

                //PG-ID
                model.setValueAt(rh.getPgid(), i,13);

            } // for end
            repaint();
        } 


        /**
        *   �󋵕\���e�[�u���̃��f��
        */
        public class StatusModel extends AbstractTableModel {
            final   int TBL_ROW = RO_MAX;
            final   int TBL_COL = 14;

            final   String[] names = {
                        "#"      , "�F" ,
                        "BtNo"   , "Proc"   , "Status" , "IP_Address", "���l" ,
                        "Status" , "IP_Address", "���l" ,
                        "P-Time" , "H-Time" , 
                        "�i��"    , "PG-ID"
            };

            private Object data[][];

            StatusModel(){
                super();

                data = new Object[TBL_ROW][TBL_COL];

                for(int i = 0 ; i < TBL_ROW ; i++){
                    data[i][0]  = new Integer(i+1);
                    data[i][1]  = "";
                    data[i][2]  = "";
                    data[i][3]  = "";
                    data[i][4]  = "";
                    data[i][5]  = "";
                    data[i][6]  = "";
                    data[i][7]  = "";
                    data[i][8]  = "";
                    data[i][9]  = "";
                    data[i][10] = "";
                    data[i][11] = "";
                    data[i][12] = "";
                    data[i][13] = "";
                }
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
        } // public class StatusModel extends AbstractTableModel


        /**
        *
        */
        class ColorRender extends DefaultTableCellRenderer {

            ColorRender(){
                super();
            }

            public Component getTableCellRendererComponent( JTable table,
                                                        Object value,
                                                        boolean isSelected,
                                                        boolean hasFocus,
                                                        int row,int column){

                String s = "";

                if(String.class == value.getClass()) s = (String)value; 

                if(ColorString.class == value.getClass()){
                    ColorString cl = (ColorString)value;
                    s = cl.getText();
                    setForeground(cl.getColor());
                }
                super.getTableCellRendererComponent(table,
                                                    s,
                                                    isSelected,
                                                    hasFocus,
                                                    row,column);
                return(this);
            }
        } //class ColorRender extends DefaultTableCellRenderer

        /**
        *
        */
        public class ColorString {
            Color color = java.awt.Color.black;
            String string = "";

            ColorString(String s,Color c){
                string = s;
                color = c;
            }

            public String getText(){
                return string;
            }

            public String toString(){
                return string;
            }

            public Color getColor(){
                return color;
            }
        } //public class ColorString
    } //public class StatusTable extends JTable

	/**
	* �^�C�g���\���p�l���N���X
	* @return LegendPanel
	*
	*/
	public class titlPanel extends JPanel {
		titlPanel(String title, Rectangle rect){
			super();
			setLayout(null);
			setBorder( new Flush3DBorder() );
			setBackground(java.awt.Color.lightGray);

			add( createLegendLabel( rect , title ) );
		}
		
		private JLabel createLegendLabel( Rectangle rect, String title ) {
			JLabel label = new JLabel( title, JLabel.CENTER );
			label.setBounds( rect );
	        label.setLocale(new Locale("ja","JP"));
	        label.setFont(new java.awt.Font("dialog", 0, 12));
			label.setForeground( java.awt.Color.black );
			return label;
		}
	}	// titlePanel

}
