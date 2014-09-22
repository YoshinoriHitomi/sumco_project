package cz;

import java.awt.Component;
import java.awt.Dimension;
import java.awt.Graphics;
import java.awt.Rectangle;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.awt.event.ComponentListener;
import java.util.Locale;
import java.util.Vector;

import javax.swing.JButton;
import javax.swing.JCheckBox;
import javax.swing.JComboBox;
import javax.swing.JDialog;
import javax.swing.JLabel;
import javax.swing.JPanel;
import javax.swing.JScrollBar;
import javax.swing.JScrollPane;
import javax.swing.JTable;
import javax.swing.ListSelectionModel;
import javax.swing.event.ListSelectionEvent;
import javax.swing.plaf.metal.MetalBorders.Flush3DBorder;
import javax.swing.table.AbstractTableModel;
import javax.swing.table.DefaultTableCellRenderer;
import javax.swing.table.DefaultTableColumnModel;
import javax.swing.table.JTableHeader;
import javax.swing.table.TableColumn;

/**
 *   ＣＣＤ生波形Window 
 * @author  (KPK Co.,Ltd.)
 * @version 1.0 (2003/04/01)
 *
 */
public class CZCMSCCDWave extends JDialog {
    private final int   REC_MAX     = 500;

    private Vector      ccd_data    = null;
    private JButton     get_button  = null;

    private JPanel      ccd_panel   = null;
    private CCDTable    ccd_table   = null;

    private WavePanel   wave_panel  = null;

    private JComboBox   x_combo     = null;
    private JComboBox   y_combo     = null;

    private JCheckBox   slice_chk   = null;
    private JCheckBox   search_chk  = null;
    private JCheckBox   k3k4_chk    = null;

    private UpdateThread updateTh   = null;
 int fff = 0;   
    //
    //
    //
    CZCMSCCDWave(){
        super();

        setTitle("ＣＣＤ生波形");
        setSize(900,800);
        setResizable(false);
        setModal(true);

        getContentPane().setLayout(null);
        // 他基地参照機能    @20131021
        if(CZSystemDefine.REFERENCE_RUN == CZSystem.getRunLevel()){
            getContentPane().setBackground(CZSystemDefine.DEFAULT_REFERENCE_BACKGROUND_COL);
        }else{
            getContentPane().setBackground(CZSystemDefine.DEFAULT_BACKGROUND_COL);
        }

        JLabel label = null;

        ccd_panel = new JPanel();
        ccd_panel.setLayout(null);
        ccd_panel.setBounds(20, 20, 850 ,230);
        ccd_panel.setBorder(new Flush3DBorder());
        ccd_panel.setBackground(java.awt.Color.lightGray);
        getContentPane().add(ccd_panel);

        ccd_table = new CCDTable();
        JTableHeader tabHead = ccd_table.getTableHeader();
        tabHead.setReorderingAllowed(false);
    
        JScrollPane panel = new JScrollPane(ccd_table);
        panel.setBounds(20, 20, 693 ,187);
        ccd_panel.add(panel);

        get_button = new JButton("波形採取");
        get_button.setBounds(730, 20, 100, 24);
        get_button.setLocale(new Locale("ja","JP"));
        get_button.setFont(new java.awt.Font("dialog", 0, 18));
        get_button.setBorder(new Flush3DBorder());
        get_button.setForeground(java.awt.Color.black);
        get_button.addActionListener(new GetButton());
        ccd_panel.add(get_button);

        JLabel lab = null;

        lab = new JLabel("Ｘ軸",JLabel.CENTER);
        lab.setBounds(730, 100, 50, 24);
        lab.setLocale(new Locale("ja","JP"));
        lab.setFont(new java.awt.Font("dialog", 0, 18));
        lab.setBorder(new Flush3DBorder());
        lab.setForeground(java.awt.Color.black);
        ccd_panel.add(lab);
    
        x_combo = new JComboBox();
        x_combo.setBounds(780, 100, 50, 24);
        x_combo.setLocale(new Locale("ja","JP"));
        x_combo.setFont(new java.awt.Font("dialog", 0, 18));
        x_combo.setForeground(java.awt.Color.black);
        x_combo.addItem("1");
        x_combo.addItem("2");
        x_combo.addItem("3");
        x_combo.addItem("4");
        x_combo.addItem("5");
        x_combo.addItem("10");
		x_combo.setFocusable(false);	/* 2007.08.22 */
        x_combo.addActionListener(new ChgXTimes());
        ccd_panel.add(x_combo);

        lab = new JLabel("Ｙ軸",JLabel.CENTER);
        lab.setBounds(730, 130, 50, 24);
        lab.setLocale(new Locale("ja","JP"));
        lab.setFont(new java.awt.Font("dialog", 0, 18));
        lab.setBorder(new Flush3DBorder());
        lab.setForeground(java.awt.Color.black);
        ccd_panel.add(lab);

        y_combo = new JComboBox();
        y_combo.setBounds(780, 130, 50, 24);
        y_combo.setLocale(new Locale("ja","JP"));
        y_combo.setFont(new java.awt.Font("dialog", 0, 18));
        y_combo.setForeground(java.awt.Color.black);
        y_combo.addItem("1");
        y_combo.addItem("2");
        y_combo.addItem("3");
        y_combo.addItem("4");
        y_combo.addItem("5");
        y_combo.addItem("10");
		y_combo.setFocusable(false);	/* 2007.08.22 */
        y_combo.addActionListener(new ChgYTimes());
        ccd_panel.add(y_combo);

        slice_chk = new JCheckBox("スライスレベル");
        slice_chk.setBounds(730, 160, 100, 15);
        slice_chk.setLocale(new Locale("ja","JP"));
        slice_chk.setFont(new java.awt.Font("dialog", 0, 10));
        slice_chk.setBorder(new Flush3DBorder());
        slice_chk.setForeground(java.awt.Color.orange);
        slice_chk.setSelected(true);
        ccd_panel.add(slice_chk);

        search_chk = new JCheckBox("サーチエリア");
        search_chk.setBounds(730, 175, 100, 15);
        search_chk.setLocale(new Locale("ja","JP"));
        search_chk.setFont(new java.awt.Font("dialog", 0, 10));
        search_chk.setBorder(new Flush3DBorder());
        search_chk.setForeground(java.awt.Color.red);
        search_chk.setSelected(true);
        ccd_panel.add(search_chk);

        k3k4_chk = new JCheckBox("K3-K4");
        k3k4_chk.setBounds(730, 190, 100, 15);
        k3k4_chk.setLocale(new Locale("ja","JP"));
        k3k4_chk.setFont(new java.awt.Font("dialog", 0, 10));
        k3k4_chk.setBorder(new Flush3DBorder());
        k3k4_chk.setForeground(java.awt.Color.cyan);
        k3k4_chk.setSelected(true);
        ccd_panel.add(k3k4_chk);

        wave_panel = new WavePanel();
        wave_panel.setBounds(20, 260, 850 ,500);
        getContentPane().add(wave_panel);   

        updateTh = new UpdateThread();
        updateTh.setPriority(Thread.MIN_PRIORITY);
        updateTh.start();

    }

    //
    //
    //
    public boolean setDefault(){

        ccd_table.clearSelection();

        for(int i = 0 ; i < REC_MAX ; i++){
            ccd_table.setValueAt("",i,1);
            ccd_table.setValueAt("",i,2);
            ccd_table.setValueAt("",i,3);
            ccd_table.setValueAt("",i,4);
            ccd_table.setValueAt("",i,5);
            ccd_table.setValueAt("",i,6);
            ccd_table.setValueAt("",i,7);
            ccd_table.setValueAt("",i,8);
            ccd_table.setValueAt("",i,9);
        }

        ccd_data = CZSystem.getCCDWave();
        if(null == ccd_data) return false;

        int size = ccd_data.size();
        if(1    >  size) return false;
        if(REC_MAX < size) size = REC_MAX;

        for(int i = 0 ; i < size ; i++){
            CZSystemCCDWave ccd = (CZSystemCCDWave)ccd_data.elementAt(i);
            if(null == ccd) continue;

            ccd_table.setValueAt(ccd.s_time,i,1);
            ccd_table.setValueAt(ccd.batch,i,2);

            String proc = CZSystem.getProcName(ccd.p_no);
            ccd_table.setValueAt(proc,i,3);

            String tm = CZSystem.timeFormat((long)ccd.p_time);
            ccd_table.setValueAt(tm,i,4);

            String sxl = new String(ccd.single + "");
            ccd_table.setValueAt(sxl,i,5);

            String kdia = new String(ccd.k_chokei + "");
            ccd_table.setValueAt(kdia,i,6);

            String hdia = new String(ccd.h_chokei + "");
            ccd_table.setValueAt(hdia,i,7);

            String vp = new String(ccd.v_keisoku + "");
            ccd_table.setValueAt(vp,i,8);

            String hp = new String(ccd.h_keisoku + "");
            ccd_table.setValueAt(hp,i,9);
        }

        x_combo.setSelectedIndex(0);
        y_combo.setSelectedIndex(0);
        wave_panel.setTimes(getXTimes(),getYTimes());

        // 他基地参照機能    @20131021
        if(CZSystemDefine.REFERENCE_RUN == CZSystem.getRunLevel()){
            get_button.setEnabled(false);
        }

        ccd_table.repaint();

        return true;
    }


    //
    //
    //
	@SuppressWarnings("unchecked")
    public void selectCCDData(){
//@@        CZSystem.log("CZCMSCCDWave","selectCCDData()");

        int list[];
        int size;

        list = ccd_table.getSelectedRows();

        Vector data = new Vector(20);
        for(int i = 0 ; i < list.length ; i++){
//@@            CZSystem.log("CZCMSCCDWave","selectCCDData List[" + i + "][" + list[i] + "]");

            size = ccd_data.size();
            if(size <= list[i]) break;
            data.addElement(ccd_data.elementAt(list[i]));
        }
        wave_panel.writeCCD(data);
    }


    //
    //
    //
    public int getXTimes(){
        int i = x_combo.getSelectedIndex();

        switch(i){
            case 0 : return 1;
            case 1 : return 2;
            case 2 : return 3;
            case 3 : return 4;
            case 4 : return 5;
            case 5 : return 10;
            default : return 1;
        }
    }


    //
    //
    //
    public int getYTimes(){
        int i = y_combo.getSelectedIndex();

        switch(i){
            case 0 : return 1;
            case 1 : return 2;
            case 2 : return 3;
            case 3 : return 4;
            case 4 : return 5;
            case 5 : return 10;
            default : return 1;
        }
    }


    /*
    *
    */
    class  ChgXTimes implements ActionListener {
        public void actionPerformed(ActionEvent e){
            wave_panel.setTimes(getXTimes(),getYTimes());
        }
    }

    /*
    *
    */
    class  ChgYTimes implements ActionListener {
        public void actionPerformed(ActionEvent e){
            wave_panel.setTimes(getXTimes(),getYTimes());
        }
    }


    /*
    *
    */
    class GetButton implements ActionListener {
        public void actionPerformed(ActionEvent ev){
            String ro = CZSystem.getRoName();

            if(CZSystem.CZOperateWaveCollect(ro)){
                CZSystem.log("CZCMSCCDWave","GetButton CZOperateWaveCollect OK !!");
                get_button.setBackground(CZSystemDefine.BUTTON_SEND_COL);
            }
            else {
                CZSystem.log("CZCMSCCDWave","GetButton CZOperateWaveCollect NG !!");
            }
        }
    }


    /*
    *
    *
    */
    class CancelButton implements ActionListener {
        public void actionPerformed(ActionEvent ev){
            setVisible(false);
        }
    }


    /*
    *
    *       ＣＣＤ実績一覧
    *
    */
    class CCDTable extends JTable {

        private CCDTblMdl model = null;

        CCDTable(){
            super();

            try{
                setName("CCDTable");
                setAutoCreateColumnsFromModel(true);
                setSelectionMode(ListSelectionModel.MULTIPLE_INTERVAL_SELECTION);

                setLocale(new Locale("ja","JP"));
                setFont(new java.awt.Font("dialog", 0, 12));
                setRowHeight(17);

                model = new CCDTblMdl();
                setModel(model);

                DefaultTableColumnModel cmdl = (DefaultTableColumnModel)getColumnModel();
                TableColumn colum = null;
                CCDTblRenderer  ren   = null;

                // No
                colum = cmdl.getColumn(0);
                colum.setMaxWidth(25);
                colum.setMinWidth(25);
                colum.setWidth(25);

                // 採取日時
                colum = cmdl.getColumn(1);
                colum.setMaxWidth(140);
                colum.setMinWidth(140);
                colum.setWidth(140);
                ren = new CCDTblRenderer();
                ren.setHorizontalAlignment(ren.CENTER);
                colum.setCellRenderer(ren);

                // バッチNo
                colum = cmdl.getColumn(2);
                colum.setMaxWidth(80);
                colum.setMinWidth(80);
                colum.setWidth(80);

                // プロセス
                colum = cmdl.getColumn(3);
                colum.setMaxWidth(60);
                colum.setMinWidth(60);
                colum.setWidth(60);

                // プロセス時間
                colum = cmdl.getColumn(4);
                colum.setMaxWidth(80);
                colum.setMinWidth(80);
                colum.setWidth(80);
                ren = new CCDTblRenderer();
                ren.setHorizontalAlignment(ren.CENTER);
                colum.setCellRenderer(ren);

                // SXL長
                colum = cmdl.getColumn(5);
                colum.setMaxWidth(60);
                colum.setMinWidth(60);
                colum.setWidth(60);
                ren = new CCDTblRenderer();
                ren.setHorizontalAlignment(ren.RIGHT);
                colum.setCellRenderer(ren);

                // 計算直径
                colum = cmdl.getColumn(6);
                colum.setMaxWidth(60);
                colum.setMinWidth(60);
                colum.setWidth(60);
                ren = new CCDTblRenderer();
                ren.setHorizontalAlignment(ren.RIGHT);
                colum.setCellRenderer(ren);

                // 平均直径
                colum = cmdl.getColumn(7);
                colum.setMaxWidth(60);
                colum.setMinWidth(60);
                colum.setWidth(60);
                ren = new CCDTblRenderer();
                ren.setHorizontalAlignment(ren.RIGHT);
                colum.setCellRenderer(ren);

                // 垂直位置
                colum = cmdl.getColumn(8);
                colum.setMaxWidth(55);
                colum.setMinWidth(55);
                colum.setWidth(55);
                ren = new CCDTblRenderer();
                ren.setHorizontalAlignment(ren.RIGHT);
                colum.setCellRenderer(ren);

                // 水平位置
                colum = cmdl.getColumn(9);
                colum.setMaxWidth(55);
                colum.setMinWidth(55);
                colum.setWidth(55);
                ren = new CCDTblRenderer();
                ren.setHorizontalAlignment(ren.RIGHT);
                colum.setCellRenderer(ren);
            }
            catch (Throwable e) {
                CZSystem.handleException(e);
            }
        }

        //
        //
        //
        public void valueChanged(ListSelectionEvent e){
            super.valueChanged(e);
            if(e.getValueIsAdjusting()) return;

            if(0 > getSelectedRow()) return;
            selectCCDData();
        }

        //
        //
        //
        public void setData(int gr,int tbl){

//@@            CZSystem.log("CZCMSCCDWave CCDTable","setData [" + gr + "][" + tbl + "]");

        }
    }

    /*
    *
    *       ＣＣＤ実績一覧：モデル
    *
    */
    public class CCDTblMdl extends AbstractTableModel {

        private int TBL_ROW     = REC_MAX;
        final   int TBL_COL     = 10;

        final String[] names = {" # "          , "採取日時" ,   
                                "Bt"           , "プロセス" ,
                                "プロセス時間" , "SxL長",
                                "計算直径"     , "平均直径" ,
                                "垂直位置"     , "水平位置" };

        private Object  data[][];

        CCDTblMdl(){
            super();

            data = new Object[TBL_ROW][TBL_COL];

            String empty   = new String("");

            for(int i = 0 ; i < TBL_ROW ; i++){
                data[i][0] = new Integer(i+1);
                data[i][1] = empty;
                data[i][2] = empty;
                data[i][3] = empty;
                data[i][4] = empty;
                data[i][5] = empty;
                data[i][6] = empty;
                data[i][7] = empty;
                data[i][8] = empty;
                data[i][9] = empty;
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
    }


    /*
    *
    *       ＣＣＤ実績一覧：レンダラー
    *
    */
    class CCDTblRenderer extends DefaultTableCellRenderer {

        CCDTblRenderer(){
            super();
            setLocale(new Locale("ja","JP"));
            setFont(new java.awt.Font("dialog", 0, 12));
        }

        public Component getTableCellRendererComponent( JTable table,
                                                    Object value,
                                                    boolean isSelected,
                                                    boolean hasFocus,
                                                    int row,int column){


            super.getTableCellRendererComponent(table,
                                                value,
                                                isSelected,
                                                hasFocus,
                                                row,column);
            return(this);
        }
    }


    /*
    *
    *       ＣＣＤ波形表示
    *
    */
    class WavePanel extends JPanel implements ComponentListener {
        private YView       y_view  = null; 
        private Rectangle   y_rec   = null;
        private int         y_pixel = 260;
        private int         y_inc   = 20;
        private int         y_times = 1;

        private XView       x_view  = null; 
        private Rectangle   x_rec   = null;
        private int         x_pixel = 2100;
        private int         x_inc   = 100;
        private int         x_times = 1;

        private MainView    main_view = null;   
        private Rectangle   main_rec  = null;
        private JScrollPane main_sc   = null;


        WavePanel(){
            super();
            setLayout(null);
            setBorder(new Flush3DBorder());
            setBackground(java.awt.Color.lightGray);

            // Y
            y_view = new YView();
            y_view.addComponentListener(this);
            JScrollPane panel = new JScrollPane(y_view);
            panel.setBounds(780, 20, 50 ,420);
            panel.setVerticalScrollBarPolicy(JScrollPane.VERTICAL_SCROLLBAR_ALWAYS);
            panel.setHorizontalScrollBarPolicy(JScrollPane.HORIZONTAL_SCROLLBAR_NEVER);		//2007-05-15 INAS
            add(panel);
            y_rec = panel.getViewportBorderBounds();
            y_view.setPixel(y_rec,y_pixel,y_inc,y_times);
            
//@@            CZSystem.log("CZCMSCCDWave WavePanel",
//@@                "YREC width[" + y_rec.width + "] height[" + y_rec.height + "]");

            // X
            x_view = new XView();
            x_view.addComponentListener(this);
            panel = new JScrollPane(x_view);
            panel.setBounds(20, 440, 760 ,40);
            panel.setVerticalScrollBarPolicy(JScrollPane.VERTICAL_SCROLLBAR_NEVER);			//2007-05-15 INAS
            panel.setHorizontalScrollBarPolicy(JScrollPane.HORIZONTAL_SCROLLBAR_ALWAYS);
            add(panel);
            x_rec = panel.getViewportBorderBounds();
            x_view.setPixel(x_rec,x_pixel,x_inc,x_times);
//@@            CZSystem.log("CZCMSCCDWave WavePanel",
//@@                "XREC width[" + x_rec.width + "] height[" + x_rec.height + "]");


            // MAIN
            JScrollBar bar = null;
            main_view = new MainView();
            main_sc = new JScrollPane(main_view);
            main_sc.setBounds(20, 20, 760 ,420);
            main_sc.setVerticalScrollBarPolicy(JScrollPane.VERTICAL_SCROLLBAR_NEVER);		//2007-05-15 INAS
            main_sc.setHorizontalScrollBarPolicy(JScrollPane.HORIZONTAL_SCROLLBAR_NEVER);	//2007-05-15 INAS
            add(main_sc);
            main_rec = main_sc.getViewportBorderBounds();
            main_view.setPixel(main_rec,x_pixel,x_inc,x_times,y_pixel,y_inc,y_times);
        }

        //
        //
        //
        public void writeCCD(Vector dat){
            main_view.writeCCD(dat);
        }

        //
        //
        //
        public void setTimes(int x,int y){
            x_times = x;
            y_times = y;

//@@            CZSystem.log("CZCMSCCDWave WavePanel",
//@@                "Times width[" + y_rec.height + "] height[" + y_times + "]");
            y_view.setTimes(y_times);
            x_view.setTimes(x_times);
            main_view.setTimes(x_times,y_times);

            y_view.repaint();
            x_view.repaint();
            main_view.repaint();
        }


        //
        // ComponentListener
        //
        public void componentResized(java.awt.event.ComponentEvent e){

        }

        public void componentMoved(java.awt.event.ComponentEvent e){
//@@            CZSystem.log("CZCMSCCDWave WavePanel","componentMoved");

            if(x_view == e.getComponent()){
                main_view.setLocation(x_view.getX(),main_view.getY());
                main_view.repaint();    
            }

            if(y_view == e.getComponent()){
                main_view.setLocation(main_view.getX(),y_view.getY());
                main_view.repaint();    
            }
        }

        public void componentShown(java.awt.event.ComponentEvent e){

        }

        public void componentHidden(java.awt.event.ComponentEvent e){

        }


        /*
        *
        */
        class MainView extends JPanel {
            private Rectangle default_rec;
            private int x_pixel = 2000;
            private int y_pixel = 255;
            private int x_inc   = 100;
            private int y_inc   = 20;
            private int x_times = 1;
            private int y_times = 1;

            CZSystemCCDWave     wave[];

            MainView(){
                super();
                setName("MainView");
                setLayout(null);
                setBackground(java.awt.Color.black);
            }

            //
            //
            //
            public void setTimes(int xt,int yt){
                x_times = xt;
                y_times = yt;

                int x = default_rec.width  * x_times;
                int y = default_rec.height * y_times;
                
                setSize(new Dimension(x,y));
                setPreferredSize(new Dimension(x, y));		//2007-05-15 INAS
            }

            //
            //
            //
            public void setPixel(Rectangle d,int xp,int xi,int xt,int yp,int yi,int yt){
                default_rec = d;
                x_pixel     = xp;
                x_inc       = xi;
                x_times     = xt;

                y_pixel     = yp;
                y_inc       = yi;
                y_times     = yt;
            }


            //
            //
            //
            public void writeCCD(Vector dat){
                int size = dat.size();

                wave = new CZSystemCCDWave[size];   

                int j = 0;
                for(int i = size -1 ; -1 < i ; i--){
                    wave[j] = (CZSystemCCDWave)dat.elementAt(i);
                    j++;
                }
                repaint();
            }


            //
            //
            //
            public void paint(Graphics g){
                Dimension d = getSize(null);
//@@                CZSystem.log("CZCMSCCDWave WavePanel",
//@@                    "MainView.paint() Y width[" + d.width + "] height[" + d.height + "]");
                g.setColor(java.awt.Color.black);
                g.fillRect(0,0,d.width,d.height);

                g.setColor(java.awt.Color.gray);

                float yp     = (float)d.height / (float)y_pixel;
                float yp_inc = yp * (float)y_inc;

                int ys_inc = 0;
                for(float y = (float)d.height ; y > 0 ; y-=yp_inc){
                    g.drawLine(0,(int)y,d.width,(int)y);
                    ys_inc += y_inc;
                }

                float xp     = (float)d.width / (float)x_pixel;
                float xp_inc = xp * (float)x_inc;

                int xs_inc = 0;
                for(float x = 0.0f ; x < d.width ; x+=xp_inc){
                    g.drawLine((int)x , 0 , (int)x , d.height);
                    xs_inc += x_inc;
                }

                // 波形描画
                if(null == wave) return;
                if(1 > wave.length )return;

                g.setColor(java.awt.Color.lightGray);
                
                int wave_length = wave.length;
                int wave_last   = wave.length -1 ;
                int loop;
                int val ;
                String a = null;
                char c[];
                int y[];
                int x[];
                int size = 0;
                int len;
                int word = 2;


                for(loop = 0 ; loop < wave_length ; loop++){
                    size = wave[loop].len;
                    len  = size * word;

                    if (null == wave[loop].data) continue;  //@@
                    c = wave[loop].data.toCharArray();
                    y = new int[size];
                    x = new int[size];

                    int j = 0;
                    for(int i = 0 ; i < len ; i+=word){
                        a = new String(c,i,word);
                        val = Integer.parseInt(a,16);
    
                        y[j] = d.height - (int)(yp * (float)val);
                        x[j] = (int)(xp * (float)j);
                        j++;
                    }

                    if(loop == (wave_last)) g.setColor(java.awt.Color.green);
                    g.drawPolyline(x,y,size);
                } // for end

                // スライスレベル描画
                if(slice_chk.isSelected()){
                    g.setColor(java.awt.Color.orange);
                    word = 4;
                    len = wave[wave_last].slice.length();
                    c   = wave[wave_last].slice.toCharArray();  
                    for(int i = 0 ; i < len ; i+=word){
                        a = new String(c,i,word);
                        val = Integer.parseInt(a,16);

                        int y1 = d.height - (int)(yp * (float)val);
                        g.drawLine(0, y1, d.width, y1); 
                    }
                }

                // サーチエリア描画
                if(search_chk.isSelected()){
                    g.setColor(java.awt.Color.red);
                    word = 4;
                    len = wave[wave_last].search.length();
                    c   = wave[wave_last].search.toCharArray(); 
                    for(int i = 0 ; i < len ; i+=word){
                        try {           //@@   取敢えずエラーを無視する。
                            a = new String(c,i,word);
                            val = Integer.parseInt(a,16);

                            int x1 = (int)(xp * (float)val);
                            g.drawLine(x1, 0, x1, d.height);
                        } catch (Exception e) {
                        }
                    }
                }
                // サーチスタート、エンドアドレス
                if(k3k4_chk.isSelected()){
                    g.setColor(java.awt.Color.cyan);
                    int x1 = (int)(xp * (float)wave[wave_last].s_start);
                    g.drawLine(x1, 0, x1, d.height);    

                    x1 = (int)(xp * (float)wave[wave_last].s_end);
                    g.drawLine(x1, 0, x1, d.height);    
                }
            } //paint
        } //end MainPanel   



        /*
        *
        */
        class YView extends JPanel {
            private Rectangle default_rec;
            private int pixel  = 255;
            private int inc    = 25;
            private int times  = 1;

            YView(){
                super();
                setName("YView");
                setLayout(null);
                setBackground(java.awt.Color.black);
            }

            //
            //
            //
            public void setTimes(int t){
                times = t;
                int x = default_rec.width;
                int y = default_rec.height * times;
        
                setSize(new Dimension(x,y));
                setPreferredSize(new Dimension(x, y));		//2007-05-15 INAS
            }

            //
            //
            //
            public void setPixel(Rectangle d,int p,int i,int t){
                default_rec = d;
                pixel       = p;
                inc         = i;
                times       = t;
            }

            //
            //
            //
            public void paint(Graphics g){
                Dimension d = getSize(null);
//@@                CZSystem.log("CZCMSCCDWave","YView Y width[" + d.width + "] height[" + d.height + "]");

                g.setColor(java.awt.Color.black);
                g.fillRect(0,0,d.width,d.height);

                g.setColor(java.awt.Color.gray);

                float p     = (float)d.height / (float)pixel;
                float p_inc = p * (float)inc;

                int s_inc = 0;
                for(float y = (float)d.height ; y > 0 ; y-=p_inc){
                    g.drawLine(0,(int)y,d.width,(int)y);
                    String s = new String(s_inc + "");
                    g.drawString(s,5,(int)(y-2));
                    s_inc += inc;
                }
            }
        } //end YPanel  



        /*
        *
        */
        class XView extends JPanel {
            private Rectangle default_rec;
            private int pixel  = 2000;
            private int inc    = 100;
            private int times  = 1;

            XView(){
                super();
                setName("XView");
                setLayout(null);
                setBackground(java.awt.Color.black);
            }

            //
            //
            //
            public void setTimes(int t){
                times = t;
                int x = default_rec.width * times;
                int y = default_rec.height;
                setSize(new Dimension(x,y));
                setPreferredSize(new Dimension(x, y));		//2007-05-15 INAS
            }


            //
            //
            //
            public void setPixel(Rectangle d,int p,int i,int t){
                default_rec = d;
                pixel       = p;
                inc     = i;
                times       = t;
            }

            //
            //
            //
            public void paint(Graphics g){
                Dimension d = getSize(null);
//@@                CZSystem.log("CZCMSCCDWave","XView X width[" + d.width + "] height[" + d.height + "]");

                g.setColor(java.awt.Color.black);
                g.fillRect(0,0,d.width,d.height);

                g.setColor(java.awt.Color.gray);

                float p     = (float)d.width / (float)pixel;
                float p_inc = p * (float)inc;

                int s_inc = 0;
                for(float x = 0.0f ; x < d.width ; x+=p_inc){
                    g.drawLine((int)x , 0 , (int)x , d.width);
                    String s = new String(s_inc + "");
                    g.drawString(s,(int)x + 5,15);
                    s_inc += inc;
                }
            }
        } //end XPanel  
    } // end WavePanel



    /*
    *
    *
    *
    */
    class UpdateThread extends Thread {

        //
        //
        //
        UpdateThread(){
        }

        //
        //
        //
        public void run(){

//@@            CZSystem.log("CZCMSCCDWave UpdateThread","START");

            CZSystemQueue   que = new CZSystemQueue(5);
            CZEventAdapter  adp = new CZEventAdapter(que);
            CZEventSender.addCZEventListener(adp);

            while(true){
                try{
                    CZEventCL event = (CZEventCL)que.waitObject();

                    switch(event.getEvent()){
                        case CZEventCL.EV_1021    :
                            get_button.setBackground(CZSystemDefine.BUTTON_WAIT_COL);
                            break;
                        case CZEventCL.EV_8021    :
                            setDefault();
                            get_button.setBackground(CZSystemDefine.BUTTON_NORMAL_COL);
                            break;
                        default  :     break;
                    } // switch end
                }
                catch(Exception e){

                }
            } // while end
        }
    }
}
