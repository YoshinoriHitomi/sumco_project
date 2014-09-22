package cz;

import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.text.DecimalFormat;
import java.util.Locale;

import javax.swing.JButton;
import javax.swing.JLabel;
import javax.swing.JPanel;
import javax.swing.JScrollPane;
import javax.swing.JTable;
import javax.swing.JViewport;
import javax.swing.ListSelectionModel;
import javax.swing.plaf.metal.MetalBorders.Flush3DBorder;
import javax.swing.table.AbstractTableModel;
import javax.swing.table.DefaultTableColumnModel;
import javax.swing.table.TableColumn;

/***********************************************************
 *
 *   集中監視操作用パネル-2
 * @author  (KPK Co.,Ltd.)
 * @version 1.0 (2003/04/01)
 *
 ***********************************************************/
public class CZCMSControlPanel_2 extends JPanel implements Runnable {

    //直径
    private JButton diaButton   = null;
    //シード
    private JButton seedPosButton   = null;
    private JButton seedRPMButton   = null;
    private JButton seedSpeedButton = null;
    //ルツボ
    private JButton cruPosButton    = null;
    private JButton cruRPMButton    = null;
    private JButton cruSpeedButton  = null;
    //ヒータ
    private JButton heaPwM1Button   = null;
    private JButton heaPwM2Button   = null;
    private JButton heaPwButButton  = null;
    private JButton heaTempButButton= null;
    //実績表示
    private ValuePanel valPanel = null;

    // ---------- コンストラクタ ---------------------------
    CZCMSControlPanel_2(){
        super();

        try{
            setName("CZCMSControlPanel_2");
            setLayout(null);
            setBackground(java.awt.Color.lightGray);

            JLabel lab = null;

            //直径関係
            lab = new JLabel("直  径",JLabel.CENTER);
            lab.setBounds(20, 10, 100, 30);
            lab.setLocale(new Locale("ja","JP"));
            lab.setFont(new java.awt.Font("dialog", 0, 18));
            lab.setBorder(new Flush3DBorder());
            lab.setForeground(java.awt.Color.black);
            add(lab);

            diaButton = new JButton("000.0");
            diaButton.setBounds(120, 10, 100, 30);
            diaButton.setLocale(new Locale("ja","JP"));
            diaButton.setFont(new java.awt.Font("dialog", 0, 18));
            diaButton.setBorder(new Flush3DBorder());
            diaButton.setBackground(java.awt.Color.lightGray);
            add(diaButton);

            //４軸関係
            lab = new JLabel("シード",JLabel.CENTER);
            lab.setBounds(120, 50, 100, 30);
            lab.setLocale(new Locale("ja","JP"));
            lab.setFont(new java.awt.Font("dialog", 0, 18));
            lab.setBorder(new Flush3DBorder());
            lab.setForeground(java.awt.Color.black);
            add(lab);

            lab = new JLabel("ルツボ",JLabel.CENTER);
            lab.setBounds(220, 50, 100, 30);
            lab.setLocale(new Locale("ja","JP"));
            lab.setFont(new java.awt.Font("dialog", 0, 18));
            lab.setBorder(new Flush3DBorder());
            lab.setForeground(java.awt.Color.black);
                add(lab);

            lab = new JLabel("位  置",JLabel.CENTER);
            lab.setBounds(20, 80, 100, 30);
            lab.setLocale(new Locale("ja","JP"));
            lab.setFont(new java.awt.Font("dialog", 0, 18));
            lab.setBorder(new Flush3DBorder());
            lab.setForeground(java.awt.Color.black);
            add(lab);

            lab = new JLabel("回  転",JLabel.CENTER);
            lab.setBounds(20, 110, 100, 30);
            lab.setLocale(new Locale("ja","JP"));
            lab.setFont(new java.awt.Font("dialog", 0, 18));
            lab.setBorder(new Flush3DBorder());
            lab.setForeground(java.awt.Color.black);
            add(lab);

            lab = new JLabel("速  度",JLabel.CENTER);
            lab.setBounds(20, 140, 100, 30);
            lab.setLocale(new Locale("ja","JP"));
            lab.setFont(new java.awt.Font("dialog", 0, 18));
            lab.setBorder(new Flush3DBorder());
            lab.setForeground(java.awt.Color.black);
            add(lab);

            //シード位置
            seedPosButton = new JButton("000.0");
            seedPosButton.setBounds(120, 80, 100, 30);
            seedPosButton.setLocale(new Locale("ja","JP"));
            seedPosButton.setFont(new java.awt.Font("dialog", 0, 18));
            seedPosButton.setBorder(new Flush3DBorder());
            seedPosButton.setBackground(java.awt.Color.lightGray);
            seedPosButton.addActionListener(new CMSSeedPosition());
            add(seedPosButton);

            //シード回転
            seedRPMButton = new JButton("00.000");
            seedRPMButton.setBounds(120, 110, 100, 30);
            seedRPMButton.setLocale(new Locale("ja","JP"));
            seedRPMButton.setFont(new java.awt.Font("dialog", 0, 18));
            seedRPMButton.setBorder(new Flush3DBorder());
            seedRPMButton.setBackground(java.awt.Color.lightGray);
            seedRPMButton.addActionListener(new CMSSeedRotation());
            add(seedRPMButton);

            //シード速度
            seedSpeedButton = new JButton("00.000");
            seedSpeedButton.setBounds(120, 140, 100, 30);
            seedSpeedButton.setLocale(new Locale("ja","JP"));
            seedSpeedButton.setFont(new java.awt.Font("dialog", 0, 18));
            seedSpeedButton.setBorder(new Flush3DBorder());
            seedSpeedButton.setBackground(java.awt.Color.lightGray);
            seedSpeedButton.addActionListener(new CMSSeedSpeed());
            add(seedSpeedButton);

            //ルツボ位置
            cruPosButton = new JButton("000.0");
            cruPosButton.setBounds(220, 80, 100, 30);
            cruPosButton.setLocale(new Locale("ja","JP"));
            cruPosButton.setFont(new java.awt.Font("dialog", 0, 18));
            cruPosButton.setBorder(new Flush3DBorder());
            cruPosButton.setBackground(java.awt.Color.lightGray);
            cruPosButton.addActionListener(new CMSCruciblePosition());
            add(cruPosButton);

            //ルツボ回転
            cruRPMButton = new JButton("00.000");
            cruRPMButton.setBounds(220, 110, 100, 30);
            cruRPMButton.setLocale(new Locale("ja","JP"));
            cruRPMButton.setFont(new java.awt.Font("dialog", 0, 18));
            cruRPMButton.setBorder(new Flush3DBorder());
            cruRPMButton.setBackground(java.awt.Color.lightGray);
            cruRPMButton.addActionListener(new CMSCrucibleRotation());
            add(cruRPMButton);

            //ルツボ速度
            cruSpeedButton = new JButton("00.000");
            cruSpeedButton.setBounds(220, 140, 100, 30);
            cruSpeedButton.setLocale(new Locale("ja","JP"));
            cruSpeedButton.setFont(new java.awt.Font("dialog", 0, 18));
            cruSpeedButton.setBorder(new Flush3DBorder());
            cruSpeedButton.setBackground(java.awt.Color.lightGray);
            cruSpeedButton.addActionListener(new CMSCrucibleSpeed());
            add(cruSpeedButton);

            //
            //ヒータ関係
            lab = new JLabel("ヒーター",JLabel.CENTER);
            lab.setBounds(440, 50, 100, 30);
            lab.setLocale(new Locale("ja","JP"));
            lab.setFont(new java.awt.Font("dialog", 0, 18));
            lab.setBorder(new Flush3DBorder());
            lab.setForeground(java.awt.Color.black);
            add(lab);

            lab = new JLabel("Main-1PW",JLabel.CENTER);
            lab.setBounds(340, 80, 100, 30);
            lab.setLocale(new Locale("ja","JP"));
            lab.setFont(new java.awt.Font("dialog", 0, 14));
            lab.setBorder(new Flush3DBorder());
            lab.setForeground(java.awt.Color.black);
            add(lab);

            lab = new JLabel("Main-2PW",JLabel.CENTER);
            lab.setBounds(340, 110, 100, 30);
            lab.setLocale(new Locale("ja","JP"));
            lab.setFont(new java.awt.Font("dialog", 0, 14));
            lab.setBorder(new Flush3DBorder());
            lab.setForeground(java.awt.Color.black);
            add(lab);

            lab = new JLabel("Bottom-PW",JLabel.CENTER);
            lab.setBounds(340, 140, 100, 30);
            lab.setLocale(new Locale("ja","JP"));
            lab.setFont(new java.awt.Font("dialog", 0, 14));
            lab.setBorder(new Flush3DBorder());
            lab.setForeground(java.awt.Color.black);
            add(lab);

            lab = new JLabel("Main-1T",JLabel.CENTER);
            lab.setBounds(340, 180, 100, 30);
            lab.setLocale(new Locale("ja","JP"));
            lab.setFont(new java.awt.Font("dialog", 0, 14));
            lab.setBorder(new Flush3DBorder());
            lab.setForeground(java.awt.Color.black);
            add(lab);

            //メインヒーター１PW
            heaPwM1Button = new JButton("00.000");
            heaPwM1Button.setBounds(440, 80, 100, 30);
            heaPwM1Button.setLocale(new Locale("ja","JP"));
            heaPwM1Button.setFont(new java.awt.Font("dialog", 0, 18));
            heaPwM1Button.setBorder(new Flush3DBorder());
            heaPwM1Button.setBackground(java.awt.Color.lightGray);
            add(heaPwM1Button);

            //メインヒーター２PW
            heaPwM2Button = new JButton("00.000");
            heaPwM2Button.setBounds(440, 110, 100, 30);
            heaPwM2Button.setLocale(new Locale("ja","JP"));
            heaPwM2Button.setFont(new java.awt.Font("dialog", 0, 18));
            heaPwM2Button.setBorder(new Flush3DBorder());
            heaPwM2Button.setBackground(java.awt.Color.lightGray);
            add(heaPwM2Button);

            //ボトムヒーターPW
            heaPwButButton = new JButton("00.000");
            heaPwButButton.setBounds(440, 140, 100, 30);
            heaPwButButton.setLocale(new Locale("ja","JP"));
            heaPwButButton.setFont(new java.awt.Font("dialog", 0, 18));
            heaPwButButton.setBorder(new Flush3DBorder());
            heaPwButButton.setBackground(java.awt.Color.lightGray);
            add(heaPwButButton);

            //メインヒーター１Ｔ(温度)
            heaTempButButton = new JButton("00.000");
            heaTempButButton.setBounds(440, 180, 100, 30);
            heaTempButButton.setLocale(new Locale("ja","JP"));
            heaTempButButton.setFont(new java.awt.Font("dialog", 0, 18));
            heaTempButButton.setBorder(new Flush3DBorder());
            heaTempButButton.setBackground(java.awt.Color.lightGray);
            add(heaTempButButton);

            //実績表示
            valPanel = new ValuePanel();
            add(valPanel);

//@@            CZSystem.log("CZSetPanel CZSetPanel","new");
        }
        catch (Throwable e) {
            CZSystem.handleException(e);
        }
    }

    //
    //
    //
    public void run(){
        CZSystemQueue   que = new CZSystemQueue(20);
        CZEventAdapter  adp = new CZEventAdapter(que);
        CZEventSender.addCZEventListener(adp);

        while(true){
            try{
                CZEventCL event = (CZEventCL)que.waitObject();
//@@                CZSystem.log("CZSetPanel run","1");
                if(event.getEvent() == CZEventCL.PV_RECEIVE){
                    updatePV();
                    valPanel.alterTbl();
                }
                if(event.getEvent() == CZEventCL.RO_CHANGE){ 
                }
            }
            catch(Exception e){
            }
//@@        CZSystem.log("CZSetPanel run","2");
        } // while end
    }

    //
    //
    //
    private int updatePV(){ 
        DecimalFormat   format   = null;
        float   val;

        //直径
        val = CZPV.getPVData(25  - 1);
        format   = new DecimalFormat("000.0");
        diaButton.setText(format.format(val));

        //シード位置
        val = CZPV.getPVData(22  - 1);
        format   = new DecimalFormat("0000.0");
        seedPosButton.setText(format.format(val));

        //シード回転
        val = CZPV.getPVData(19  - 1);
        format   = new DecimalFormat("00.000");
        seedRPMButton.setText(format.format(val));

        //シード速度
        val = CZPV.getPVData(18  - 1);
        format   = new DecimalFormat("00.0000");
        seedSpeedButton.setText(format.format(val));

        //ルツボ位置
        val = CZPV.getPVData(23  - 1);
        format   = new DecimalFormat("0000.0");
        cruPosButton.setText(format.format(val));

        //ルツボ回転
        val = CZPV.getPVData(21  - 1);
        format   = new DecimalFormat("00.000");
        cruRPMButton.setText(format.format(val));

        //ルツボ速度
        val = CZPV.getPVData(20  - 1);
        format   = new DecimalFormat("00.0000");
        cruSpeedButton.setText(format.format(val));

        //メインヒーター１PW
        val = CZPV.getPVData(12  - 1);
        format   = new DecimalFormat("000.00");
        heaPwM1Button.setText(format.format(val));

        //メインヒーター２PW
        val = CZPV.getPVData(13  - 1);
        format   = new DecimalFormat("000.00");
        heaPwM2Button.setText(format.format(val));

        //ボトムヒーター
        val = CZPV.getPVData(14  - 1);
        format   = new DecimalFormat("000.00");
        heaPwButButton.setText(format.format(val));

        //メインヒーター１T
        val = CZPV.getPVData(15  - 1);
        format   = new DecimalFormat("0000.0");
        heaTempButButton.setText(format.format(val));

        return 0;
    }
    

    //
    //  リモート操作
    //
    //
    //      シード位置
    //
    class CMSSeedPosition implements ActionListener {
        private CZCMSSeedPosition obj = null;
        public void actionPerformed(ActionEvent e){
            if(null == obj) obj = new CZCMSSeedPosition();
            obj.setDefault();
            obj.setVisible(true);
        }
    }

    //
    //      シード回転
    //
    class CMSSeedRotation implements ActionListener {
        private CZCMSSeedRotation obj = null;

        public void actionPerformed(ActionEvent e){
            if(null == obj) obj = new CZCMSSeedRotation();
            obj.setDefault();
            obj.setVisible(true);
        }
    }

    //
    //      シード速度
    //
    class CMSSeedSpeed implements ActionListener {
        private CZCMSSeedSpeed obj = null;

        public void actionPerformed(ActionEvent e){
            if(null == obj) obj = new CZCMSSeedSpeed();
            obj.setDefault();
            obj.setVisible(true);
        }
    }

    //
    //      ルツボ位置
    //
    class CMSCruciblePosition implements ActionListener {
        private CZCMSCruciblePosition obj = null;

        public void actionPerformed(ActionEvent e){
            if(null == obj) obj = new CZCMSCruciblePosition();
            obj.setDefault();
            obj.setVisible(true);
        }
    }

    //
    //      ルツボ回転
    //
    class CMSCrucibleRotation implements ActionListener {
        private CZCMSCrucibleRotation obj = null;

        public void actionPerformed(ActionEvent e){
            if(null == obj) obj = new CZCMSCrucibleRotation();
            obj.setDefault();
            obj.setVisible(true);
        }
    }

    //
    //      ルツボ速度
    //
    class CMSCrucibleSpeed implements ActionListener {
        private CZCMSCrucibleSpeed obj = null;

        public void actionPerformed(ActionEvent e){
            if(null == obj) obj = new CZCMSCrucibleSpeed();
            obj.setDefault();
            obj.setVisible(true);
        }
    }


    /*******************************************************
     *
     *
     *
     *******************************************************/
    class ValuePanel extends JScrollPane {

        private ValueTbl valTbl = null;

        ValuePanel(){
            super();

            setName("ValuePanel");
            setVerticalScrollBarPolicy(JScrollPane.VERTICAL_SCROLLBAR_NEVER);
            setHorizontalScrollBarPolicy(JScrollPane.HORIZONTAL_SCROLLBAR_NEVER);
            setBounds(560, 10 , 220, 206);
            getViewport().setScrollMode(JViewport.BACKINGSTORE_SCROLL_MODE);
            valTbl = new ValueTbl();
            setViewportView(valTbl);

        }

        //
        //
        //
        public void alterTbl(){
            valTbl.alterTbl();
        }

        /***************************************************
         *
         ***************************************************/
        class ValueTbl extends JTable {

            private ValueModel model = null;

            final String[] names = {"SXL長",
                        "引上重量(計)",
                        "結晶保持荷重(実)",
                        "残液量",
                        "液温",
                        "プルAr",
                        "トップAr",
                        "炉内圧",
                        "ヒーター温度目標",
                        "予備",
                        "ヒータON時間"};

            ValueTbl(){

                super();

                setName("ValueTbl");
                setAutoCreateColumnsFromModel(true);
                setSelectionMode(ListSelectionModel.SINGLE_SELECTION);
                setLocale(new Locale("ja","JP"));
                setFont(new java.awt.Font("dialog", 0, 12));
                setRowHeight(17);

                model = new ValueModel(11);

                setModel(model);

                DefaultTableColumnModel cmdl = (DefaultTableColumnModel)getColumnModel();
                TableColumn  colum = null;

                colum = cmdl.getColumn(0);
                colum.setMaxWidth(20);
                colum.setMinWidth(20);
                colum.setWidth(20);

                colum = cmdl.getColumn(1);
                colum.setMaxWidth(110);
                colum.setMinWidth(110);
                colum.setWidth(110);

                colum = cmdl.getColumn(2);
                colum.setMinWidth(85);
                colum.setWidth(85);

                int size = names.length;
                for(int i = 0 ; i < size ; i++){
                    model.setValueAt(names[i],i,1);
                }
            }   

            //
            //
            //
            public void alterTbl(){

                if(null == model) return;

                //SXL長
                int i = 0;
                model.setValueAt(String.valueOf(CZPV.getPVData(5 - 1)),i,2);

                //引上重量(計)
                i++;
                model.setValueAt(String.valueOf(CZPV.getPVData(53 - 1)),i,2);

                //結晶保持荷重(実)
                i++;
                model.setValueAt(String.valueOf(CZPV.getPVData(113 - 1)),i,2);

                //残液量
                i++;
                model.setValueAt(String.valueOf(CZPV.getPVData(54 - 1)),i,2);

                //液温
                i++;
                model.setValueAt(String.valueOf(CZPV.getPVData(31 - 1)),i,2);

                //プルAr
                i++;
                model.setValueAt(String.valueOf(CZPV.getPVData(16 - 1)),i,2);

                //トップAr
                i++;
                model.setValueAt(String.valueOf(CZPV.getPVData(17 - 1)),i,2);

                //炉内圧
                i++;
                model.setValueAt(String.valueOf(CZPV.getPVData(33 - 1)),i,2);

                //ヒーター温度目標
                i++;
                model.setValueAt(String.valueOf(CZPV.getPVData(66 - 1)),i,2);

                //予備
                i++;

                //ヒーターON時間
                i++;
                model.setValueAt(CZSystem.timeFormat(CZPV.getHtOnTm()),i,2);

                repaint();
            }


            /***********************************************
             *
             ***********************************************/
            class ValueModel extends AbstractTableModel {
                final int TBL_COL       = 3;
                private int TBL_ROW     = 0;

                private Object data[][];

                final String[] hed = {"#","Name","Data"};

                ValueModel(int row){
                    super();

                    TBL_ROW = row;
                        data = new Object[TBL_ROW][TBL_COL];

                    int i;
                    for(i = 0 ; i < TBL_ROW ; i++){
                        data[i][0] = new Integer(i+1);
                        data[i][1] = new String("#########");
                        data[i][2] = new String("----------");
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
                    return hed[column];
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
            } // ValueModel
        } // ValueTbl 
    } // ValuePanel
}
