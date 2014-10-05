package cz;

import java.awt.Color;
import java.awt.Component;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.text.DecimalFormat;
import java.util.Locale;

import javax.swing.JButton;
import javax.swing.JDialog;
import javax.swing.JLabel;
import javax.swing.JScrollPane;
import javax.swing.JTable;
import javax.swing.ListSelectionModel;
import javax.swing.plaf.metal.MetalBorders.Flush3DBorder;
import javax.swing.table.AbstractTableModel;
import javax.swing.table.DefaultTableCellRenderer;
import javax.swing.table.DefaultTableColumnModel;
import javax.swing.table.TableColumn;

import czclass.CZNativeRoState;

/**
 *   稼働状況表示画面
 * @author  (KPK Co.,Ltd.)
 * @version 1.0 (2003/04/01)
 */
public class CZOperationalStatus extends JDialog {

    private final int   RO_MAX = 99;
//    private final int   RO_MAX = 50;

    private final int AUTO      = 2;    // 3    制御モード
    private final int PSXL      = 4;    // 5    プロセスＳＸＬ長

    private final int HT_P1     = 11;   // 12   メインヒーター１電力
    private final int HT_P2     = 12;   // 13   メインヒーター２電力
    private final int HT_PB     = 13;   // 14   ボトムヒーター電力

    private final int MAIN1_H_T = 14;   // 15   メインヒーター１温度

    private final int DIA       = 24;   // 25   直径
    private final int SXL_ST    = 17;   // 18   シード速度
    private final int SXL_RT    = 18;   // 19   シード回転
    private final int CRU_ST    = 19;   // 20   ルツボ速度
    private final int CRU_RT    = 20;   // 21   ルツボ回転

    private final int MEL_T     = 30;   // 31   液温

    private JLabel      time_lab        = null; //更新日時表示
    private JButton     send_button     = null;
    private JButton     cancel_button   = null;

    private JScrollPane st_scpanel      = null; //スクロール
    private StatusTable st_table        = null; //稼動状況テーブル

    private CZNativeRoState[] status_list;

    private UpdateThread    updateTh    = null; //更新スレッド

    /**
    *   コンストラクタ
    */
    CZOperationalStatus(){
        super();

        setTitle("稼働状況");
        setSize(1150,700);
        setResizable(false);
        setModal(false);

        getContentPane().setLayout(null);
        // 他基地参照機能    @20131021
        if(CZSystemDefine.REFERENCE_RUN == CZSystem.getRunLevel()){
            getContentPane().setBackground(CZSystemDefine.DEFAULT_REFERENCE_BACKGROUND_COL);
        }else{
            getContentPane().setBackground(CZSystemDefine.DEFAULT_BACKGROUND_COL);
        }

        JLabel lab = new JLabel("更新日時",JLabel.CENTER);
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

        //
        st_table = new StatusTable();
        st_scpanel = new JScrollPane(st_table);
        st_scpanel.setBounds(20, 50, 1110, 565);
        st_scpanel.setBorder(new Flush3DBorder());
        getContentPane().add(st_scpanel);

        updateTh = new UpdateThread();
        updateTh.setPriority(Thread.MIN_PRIORITY);
        updateTh.start();
        
        CZSystem.log("CZOperationalStatus","稼動状況　表示");
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
        status_list = CZSystem.CZNativeRoStateGet();

        if(null == status_list) return false;

        int size = status_list.length;
        if(1 > size) return false;

        st_table.setData();

        String tm = CZSystem.getDateTime();
        time_lab.setText(tm);

        return true;
    } 


    /**
    *
    */
    class SendButton implements ActionListener {
        public void actionPerformed(ActionEvent ev){
            updateStatus();
            CZSystem.sleep(5000);
        }
    }


    /**
    *
    */
    class CancelButton implements ActionListener {
        public void actionPerformed(ActionEvent ev){
            setDefault();
            setVisible(false);
        }
    }


    /**
    *   更新スレッド
    */
    class UpdateThread extends Thread {

        UpdateThread(){

        }

        public void run(){
//@@            CZSystem.log("CZOperationalStatus","UpdateThread START");

            while(true){
                updateStatus();
                CZSystem.sleep(10000);
            } // while end
        }
    }


    /**
    *   状況表示用テーブル
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

            //炉
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

            //プロセス
            ren = new ColorRender();
            colum = cmdl.getColumn(3);
            colum.setMaxWidth(60);
            colum.setMinWidth(60);
            colum.setWidth(60);
            colum.setCellRenderer(ren);

            //モード
            ren = new ColorRender();
            ren.setHorizontalAlignment(ren.CENTER);
            colum = cmdl.getColumn(4);
            colum.setMaxWidth(40);
            colum.setMinWidth(40);
            colum.setWidth(40);
            colum.setCellRenderer(ren);

            //操作炉前
            ren = new ColorRender();
            ren.setHorizontalAlignment(ren.CENTER);
            colum = cmdl.getColumn(5);
            colum.setMaxWidth(20);
            colum.setMinWidth(20);
            colum.setWidth(20);
            colum.setCellRenderer(ren);

            //操作集中監視
            ren = new ColorRender();
            ren.setHorizontalAlignment(ren.CENTER);
            colum = cmdl.getColumn(6);
            colum.setMaxWidth(20);
            colum.setMinWidth(20);
            colum.setWidth(20);
            colum.setCellRenderer(ren);


            //プロセス時間
            ren = new ColorRender();
            ren.setHorizontalAlignment(ren.CENTER);
            colum = cmdl.getColumn(7);
            colum.setMaxWidth(80);
            colum.setMinWidth(80);
            colum.setWidth(80);
            colum.setCellRenderer(ren);

            //ヒーターオン時間
            ren = new ColorRender();
            ren.setHorizontalAlignment(ren.CENTER);
            colum = cmdl.getColumn(8);
            colum.setMaxWidth(80);
            colum.setMinWidth(80);
            colum.setWidth(80);
            colum.setCellRenderer(ren);

            //DIA
            ren = new ColorRender();
            ren.setHorizontalAlignment(ren.RIGHT);
            colum = cmdl.getColumn(9);
            colum.setMaxWidth(50);
            colum.setMinWidth(50);
            colum.setWidth(50);
            colum.setCellRenderer(ren);

            //L
            ren = new ColorRender();
            ren.setHorizontalAlignment(ren.RIGHT);
            colum = cmdl.getColumn(10);
            colum.setMaxWidth(50);
            colum.setMinWidth(50);
            colum.setWidth(50);
            colum.setCellRenderer(ren);

            //SXL.ST
            ren = new ColorRender();
            ren.setHorizontalAlignment(ren.RIGHT);
            colum = cmdl.getColumn(11);
            colum.setMaxWidth(60);
            colum.setMinWidth(60);
            colum.setWidth(60);
            colum.setCellRenderer(ren);

            //SXL.RT
            ren = new ColorRender();
            ren.setHorizontalAlignment(ren.RIGHT);
            colum = cmdl.getColumn(12);
            colum.setMaxWidth(60);
            colum.setMinWidth(60);
            colum.setWidth(60);
            colum.setCellRenderer(ren);

            //CRU.ST
            ren = new ColorRender();
            ren.setHorizontalAlignment(ren.RIGHT);
            colum = cmdl.getColumn(13);
            colum.setMaxWidth(60);
            colum.setMinWidth(60);
            colum.setWidth(60);
            colum.setCellRenderer(ren);

            //CRU.RT
            ren = new ColorRender();
            ren.setHorizontalAlignment(ren.RIGHT);
            colum = cmdl.getColumn(14);
            colum.setMaxWidth(60);
            colum.setMinWidth(60);
            colum.setWidth(60);
            colum.setCellRenderer(ren);

            //MEL.T
            ren = new ColorRender();
            ren.setHorizontalAlignment(ren.RIGHT);
            colum = cmdl.getColumn(15);
            colum.setMaxWidth(50);
            colum.setMinWidth(50);
            colum.setWidth(50);
            colum.setCellRenderer(ren);

            //HEA.T1
            ren = new ColorRender();
            ren.setHorizontalAlignment(ren.RIGHT);
            colum = cmdl.getColumn(16);
            colum.setMaxWidth(50);
            colum.setMinWidth(50);
            colum.setWidth(50);
            colum.setCellRenderer(ren);

            //HEA.P1
            ren = new ColorRender();
            ren.setHorizontalAlignment(ren.RIGHT);
            colum = cmdl.getColumn(17);
            colum.setMaxWidth(50);
            colum.setMinWidth(50);
            colum.setWidth(50);
            colum.setCellRenderer(ren);

            //HEA.P2
            ren = new ColorRender();
            ren.setHorizontalAlignment(ren.RIGHT);
            colum = cmdl.getColumn(18);
            colum.setMaxWidth(50);
            colum.setMinWidth(50);
            colum.setWidth(50);
            colum.setCellRenderer(ren);

            //HEA.PB
            ren = new ColorRender();
            ren.setHorizontalAlignment(ren.RIGHT);
            colum = cmdl.getColumn(19);
            colum.setMaxWidth(50);
            colum.setMinWidth(50);
            colum.setWidth(50);
            colum.setCellRenderer(ren);
        }

        //
        //
        //
        public void setData(){

            if(null == status_list) return;
            int size = status_list.length;
            if(1 > size) return;

            float[] pv;
            DecimalFormat f1 = new DecimalFormat("##0.0");
            DecimalFormat f2 = new DecimalFormat("###0.0");
            DecimalFormat f3 = new DecimalFormat("#0.0000");

            ColorRender cell;

            for(int i = 0 ; i < size ; i++){
                CZNativeRoState st = status_list[i];

                //PVデータ
                pv = st.getData();

                //炉番
				String s = CZSystem.RoKetaChg(st.getRoName());	// 20050725 炉：表示桁数変更
                model.setValueAt(s, i,1);
//                model.setValueAt(st.getRoName(), i,1);

                //バッチNo
                model.setValueAt(st.getBatch(), i,2);

                //プロセス
                model.setValueAt(new ColorString(CZSystem.getProcName(st.getP_no()),java.awt.Color.blue), i,3);
                if(st.getDown()) model.setValueAt(new ColorString("DOWN",java.awt.Color.red), i,3);

                //モード
                cell = (ColorRender)getCellRenderer(i,4);
                int mode = (int)pv[AUTO];
                switch(mode){
                    case CZSystemDefine.PROC_MANUAL :
                        model.setValueAt(new ColorString(
                        CZSystemDefine.PROC_MODE[CZSystemDefine.PROC_MANUAL],java.awt.Color.red),i,4);
                        break;

                    case CZSystemDefine.PROC_AUTO :
                        model.setValueAt(new ColorString(
                        CZSystemDefine.PROC_MODE[CZSystemDefine.PROC_AUTO],java.awt.Color.blue),i,4);
                        break;

                    default : 
                    model.setValueAt(new ColorString("不  明",java.awt.Color.black),i,4);
                        break;
                }

                //操作
                if(st.getFrontOperate()){
                    model.setValueAt(new ColorString("操",java.awt.Color.red),i,5);
                } else {
                    model.setValueAt("",i,5);
                }

                if(st.getRemoteOperate()){
                    model.setValueAt(new ColorString("操",java.awt.Color.blue),i,6);
                } else {
                    model.setValueAt("",i,6);
                }

                //プロセス時間
                model.setValueAt(CZSystem.timeFormat(st.getP_time()), i,7);

                //ヒータオン時間
                if(0 >= st.getH_ontime()) {
                    model.setValueAt(
                        new ColorString(CZSystem.timeFormat(st.getH_ontime()),java.awt.Color.red), i,8);
                } else {
                    model.setValueAt(
                        new ColorString(CZSystem.timeFormat(st.getH_ontime()),java.awt.Color.blue), i,8);
                }

                //DIA
                model.setValueAt(f1.format(pv[DIA]),i,9);

                //L
                model.setValueAt(f2.format(pv[PSXL]),i,10);

                //SXL.ST
                model.setValueAt(f3.format(pv[SXL_ST]),i,11);
                //SXL.RT
                model.setValueAt(f3.format(pv[SXL_RT]),i,12);

                //CRU.ST
                model.setValueAt(f3.format(pv[CRU_ST]),i,13);
                //CRU.RT
                model.setValueAt(f3.format(pv[CRU_RT]),i,14);

                //MEL.T
                model.setValueAt(f2.format(pv[MEL_T]),i,15);

                //HEA.T1
                model.setValueAt(f2.format(pv[MAIN1_H_T]),i,16);

                //HEA.P1
                model.setValueAt(f1.format(pv[HT_P1]),i,17);
                //HEA.P2
                model.setValueAt(f1.format(pv[HT_P2]),i,18);
                //HEA.PB
                model.setValueAt(f1.format(pv[HT_PB]),i,19);

                //バージョン
                model.setValueAt(st.getVersion(), i,20);
            } // for end
            repaint();
        } 


        /**
        *   状況表示テーブルのモデル
        */
        public class StatusModel extends AbstractTableModel {
            final   int TBL_ROW = RO_MAX;
            final   int TBL_COL = 21;

            final   String[] names = {
                        "#"      , "炉" ,
                        "BtNo"   , "Proc"   , "Mode" , "炉", "集" ,
                        "P-Time" , "H-Time" , 
                        "DIA"    , "L" , 
                        "SXL.ST" , "SXL.RT" ,
                        "CRU.ST" , "CRU.RT" ,
                        "MEL.T"  , "HEA.T1" ,
                        "HEA.P1" , "HEA.P2" , "HEA.PB",
                        "Ver"
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
                    data[i][14] = "";
                    data[i][15] = "";
                    data[i][16] = "";
                    data[i][17] = "";
                    data[i][18] = "";
                    data[i][19] = "";
                    data[i][20] = "";
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
}
