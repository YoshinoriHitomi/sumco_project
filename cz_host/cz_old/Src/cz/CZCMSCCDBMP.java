package cz;

import java.awt.Component;
import java.awt.Dimension;
import java.awt.Graphics;
import java.awt.Image;
import java.awt.MediaTracker;
import java.awt.Rectangle;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.awt.event.ComponentListener;
import java.awt.event.MouseEvent;
import java.awt.event.MouseMotionListener;
import java.awt.image.MemoryImageSource;
import java.io.File;
import java.io.FileInputStream;
import java.util.Locale;
import java.util.Properties;
import java.util.Vector;

import javax.swing.JButton;
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
 *  ＣＣＤ画像Window 
 * @author  (KPK Co.,Ltd.)
 * @version 1.0 (2003/04/01)
 *
 */
public class CZCMSCCDBMP extends JDialog {
    private String       save_dir   = null;

    private final int    REC_MAX    = 500;

    private Vector       ccd_data   = null;
    private JButton      get_button = null;

    private JPanel       ccd_panel  = null;
    private CCDTable     ccd_table  = null;

    private WavePanel    wave_panel = null;

    private UpdateThread updateTh   = null;
    //
    //
    //
    CZCMSCCDBMP(){
        super();

        setTitle("ＣＣＤ画像");
        setSize(900,800);
        setResizable(false);
        setModal(true);

        try{
            Properties prop =  new Properties();
            FileInputStream pros = new FileInputStream(CZSystemDefine.PROPERTY_FILE);
            prop.load(pros);

            save_dir = prop.getProperty("CCD_BMP_FILE_DIR");
            if(null == save_dir) CZSystem.exit(-1,"CZCMSCCDBMP NO Propertie File null");
            if(1 > save_dir.length()) CZSystem.exit(-1,"CZCMSCCDBMP NO Propertie File name");
        }
        catch( Exception e){
            CZSystem.exit(-1,"CZCMSCCDBMP NO Propertie File");
        }

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

        get_button = new JButton("画像採取");
        get_button.setBounds(730, 20, 100, 24);
        get_button.setLocale(new Locale("ja","JP"));
        get_button.setFont(new java.awt.Font("dialog", 0, 18));
        get_button.setBorder(new Flush3DBorder());
        get_button.setForeground(java.awt.Color.black);
        get_button.addActionListener(new GetButton());
        ccd_panel.add(get_button);

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
        }

        ccd_data = CZSystem.getCCDBMP();
        if(null == ccd_data) return false;

        int size = ccd_data.size();
        if(1    >  size) return false;
        if(REC_MAX < size) size = REC_MAX;

        for(int i = 0 ; i < size ; i++){
            CZSystemCCDBMP ccd = (CZSystemCCDBMP)ccd_data.elementAt(i);
            if(null == ccd) continue;

            ccd_table.setValueAt(ccd.s_time,i,1);
            ccd_table.setValueAt(ccd.batch,i,2);

            String proc = CZSystem.getProcName(ccd.p_no);
            ccd_table.setValueAt(proc,i,3);

            String tm = CZSystem.timeFormat((long)ccd.p_time);
            ccd_table.setValueAt(tm,i,4);
            ccd_table.setValueAt(ccd.f_name,i,5);
        }

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
    public void selectCCDData(){
//@@        CZSystem.log("CZCMSCCDBMP","selectCCDData()");

        int no;
        int size = ccd_data.size();
        if(1 > size) return;

        no = ccd_table.getSelectedRow();
        if(0 > no) return;

        if(no >= size) return;

        CZSystemCCDBMP ccd = (CZSystemCCDBMP)ccd_data.elementAt(no);

//@@        CZSystem.log("CZCMSCCDBMP","selectCCDData no[" + no + "] File[" + ccd.f_name + "]");

        wave_panel.writeCCD(save_dir,ccd.f_name.trim());

        return;
    }


    /*
    *
    *
    *
    */
    class GetButton implements ActionListener {
        public void actionPerformed(ActionEvent ev){
            String ro = CZSystem.getRoName();

            if(CZSystem.CZOperateCcdCamera(ro)){
                CZSystem.log("CZCMSCCDBMP","GetButton CZOperateCcdCamera OK !!");
                get_button.setBackground(CZSystemDefine.BUTTON_SEND_COL);
            }
            else {
                CZSystem.log("CZCMSCCDBMP","GetButton CZOperateCcdCamera NG !!");
            }
        }
    }


    /*
    *
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
    *       ＣＣＤ画像実績一覧
    *
    */
    class CCDTable extends JTable {

        private CCDTblMdl model = null;

        CCDTable(){
            super();

            try{
                setName("CCDTable");
                setAutoCreateColumnsFromModel(true);
                setSelectionMode(ListSelectionModel.SINGLE_SELECTION);

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
                colum.setMaxWidth(30);
                colum.setMinWidth(30);
                colum.setWidth(30);

                // 採取日時
                colum = cmdl.getColumn(1);
                colum.setMaxWidth(160);
                colum.setMinWidth(160);
                colum.setWidth(160);
                ren = new CCDTblRenderer();
                ren.setHorizontalAlignment(ren.CENTER);
                colum.setCellRenderer(ren);

                // バッチNo
                colum = cmdl.getColumn(2);
                colum.setMaxWidth(100);
                colum.setMinWidth(100);
                colum.setWidth(100);

                // プロセス
                colum = cmdl.getColumn(3);
                colum.setMaxWidth(70);
                colum.setMinWidth(70);
                colum.setWidth(70);

                // プロセス時間
                colum = cmdl.getColumn(4);
                colum.setMaxWidth(80);
                colum.setMinWidth(80);
                colum.setWidth(80);
                ren = new CCDTblRenderer();
                ren.setHorizontalAlignment(ren.CENTER);
                colum.setCellRenderer(ren);

                // ファイル名
                colum = cmdl.getColumn(5);
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

//@@            CZSystem.log("CZCMSCCDBMP","CCDTable setData() [" + gr + "][" + tbl + "]");
        }
    }

    /*
    *
    *       ＣＣＤ実績一覧：モデル
    *
    */
    public class CCDTblMdl extends AbstractTableModel {

        private int TBL_ROW     = REC_MAX;
        final   int TBL_COL     = 6;

        final String[] names = {" # "          , "採取日時" ,   
                                "Bt"           , "プロセス" ,
                                "プロセス時間" , "ファイル名" };

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
    *       ＣＣＤ画像表示
    *
    */
    class WavePanel extends JPanel implements ComponentListener {
        private YView       y_view  = null; 
        private Rectangle   y_rec   = null;
        private int     y_pixel = 100;
        private int     y_inc   = 20;
        private int     y_times = 1;

        private XView       x_view  = null; 
        private Rectangle   x_rec   = null;
        private int     x_pixel = 100;
        private int     x_inc   = 100;
        private int     x_times = 1;

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
            y_view.setPixel(y_rec);

            // X
            x_view = new XView();
            x_view.addComponentListener(this);
            panel = new JScrollPane(x_view);
            panel.setBounds(20, 440, 760 ,40);
            panel.setVerticalScrollBarPolicy(JScrollPane.VERTICAL_SCROLLBAR_NEVER);			//2007-05-15 INAS
            panel.setHorizontalScrollBarPolicy(JScrollPane.HORIZONTAL_SCROLLBAR_ALWAYS);
            add(panel);
            x_rec = panel.getViewportBorderBounds();
            x_view.setPixel(x_rec);

            // MAIN
            JScrollBar bar = null;
            main_view = new MainView();
            main_sc = new JScrollPane(main_view);
            main_sc.setBounds(20, 20, 760 ,420);
            main_sc.setVerticalScrollBarPolicy(JScrollPane.VERTICAL_SCROLLBAR_NEVER);		//2007-05-15 INAS
            main_sc.setHorizontalScrollBarPolicy(JScrollPane.HORIZONTAL_SCROLLBAR_NEVER);	//2007-05-15 INAS
            add(main_sc);
            main_rec = main_sc.getViewportBorderBounds();
            main_view.setPixel(main_rec);

        }

        //
        //
        //
        public void writeCCD(String path,String file_name){
//@@            CZSystem.log("CZCMSCCDBMP","WavePanel writeCCD() Path[" + path + "] File[" + file_name + "]");
            if((null == path) || (null == file_name)) return;

            File file = new File(path, file_name);
            if(!file.exists()){
                CZSystem.log("CZCMSCCDBMP","WavePanel writeCCD() Not found File [" + path + file_name + "]");
                return;
            }

            CZBMPReader reader;
            int[]   pix;
            Image   img = null;
            int wid = 200, hei = 200;

            try {
                reader = new CZBMPReader(file);
                reader.read();

                pix = reader.getPix();
                wid = reader.getWidth();
                hei = reader.getHeight();
//@@                CZSystem.log("CZCMSCCDBMP","WavePanel writeCCD() wid[" + wid + "] hei[" + hei + "]");
                img = createImage( new MemoryImageSource( wid,hei, pix, 0,wid));
            }
            catch(Exception e){
                CZSystem.log("CZCMSCCDBMP","WavePanel writeCCD() Exception[" + e + "]");
            }
            finally {

            }

            x_view.setWidth(wid);
            y_view.setHeight(hei);
            main_view.writeCCD(wid,hei,img);

        }


        //
        // ComponentListener
        //
        public void componentResized(java.awt.event.ComponentEvent e){
        }

        public void componentMoved(java.awt.event.ComponentEvent e){
//@@            CZSystem.log("CZCMSCCDBMP","WavePanel componentMoved");

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
            Image   image   = null;

            MainViewMouseMotion mouse_motion;

            MainView(){
                super();
                setName("MainView");
                setLayout(null);
                setBackground(java.awt.Color.black);

                mouse_motion = new MainViewMouseMotion();

                addMouseMotionListener(mouse_motion);
            }

            //
            //
            //
            public void setPixel(Rectangle d){
                default_rec = d;
            }


            //
            //
            //
            public void writeCCD(int wid,int hei,Image img){

                setSize(new Dimension(wid,hei));
                setPreferredSize(new Dimension(wid,hei));		//2007-05-15 INAS
                image = img;

                repaint();
            }

            //
            //
            //
            public void paint(Graphics g){
                Dimension d = getSize(null);
//@@                CZSystem.log("CZCMSCCDBMP","WavePanel MainView Y width[" + d.width + "] height[" + d.height + "]");

                g.setColor(java.awt.Color.gray);
                g.fillRect(0,0,d.width,d.height);


                if(null == image) return;

                MediaTracker mt = new MediaTracker(this);
                mt.addImage(image,0);
                try {
                    mt.waitForAll();
                }   
                catch (InterruptedException e) {
                    CZSystem.log("CZCMSCCDBMP","WavePanel MainView DMediaTracker [" + e + "]");
                }
                mouse_motion.setDefault();

//@@                CZSystem.log("CZCMSCCDBMP","WavePanel MainView Draw Image Width[" +
//@@                        image.getWidth(this) + " Height[" + image.getHeight(this) + "]");
                g.drawImage(image,0,0,this);
            } //paint

            /*
            *
            *
            */
            class MainViewMouseMotion implements MouseMotionListener {

                int old_x = -1;
                int old_y = -1;
                String  old_string = "(-1,-1)";

                //
                //
                //
                public void mouseDragged(MouseEvent e){

                }

                //
                //
                //
                public void setDefault(){
                    old_x = -1;
                    old_y = -1;
                }

                //
                //
                //
                public void mouseMoved(MouseEvent e){
                    int x = e.getX();
                    int y = e.getY();

                    Dimension d = getSize(null);

                    JPanel c = (JPanel)e.getSource();
                    Graphics g = c.getGraphics();

                    g.setXORMode(java.awt.Color.red);
                    g.setColor(java.awt.Color.white);

                    g.drawLine(old_x, 0, old_x, d.height);
                    g.drawLine(0, old_y, d.width, old_y);
                    g.drawString(old_string,old_x+25,old_y-20);


                    g.drawLine(x, 0, x, d.height);
                    g.drawLine(0, y, d.width, y);

                    String s = new String("(" + x + "," + y + ")");
                    g.drawString(s,x+25,y-20);

                    old_x = x;
                    old_y = y;
                    old_string = s;
                }
            } // MouseMotion
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
            public void setHeight(int h){
                int x = default_rec.width;
                setSize(new Dimension(x,h));
                setPreferredSize(new Dimension(x,h));		//2007-05-15 INAS
            }


            //
            //
            //
            public void setPixel(Rectangle d){
                default_rec = d;
            }

            //
            //
            //
            public void paint(Graphics g){
                Dimension d = getSize(null);
//@@                CZSystem.log("CZCMSCCDBMP","WavePanel YView Y width[" + d.width + "] height[" + d.height + "]");

                g.setColor(java.awt.Color.black);
                g.fillRect(0,0,d.width,d.height);

                g.setColor(java.awt.Color.gray);


            
                for(int i = 0 ; i < d.height ; i +=inc){
                    g.drawLine(0,i,d.width,i);
                    String s = new String(i + "");
                    g.drawString(s,5,i-2);
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
            public void setWidth(int w){
                int y = default_rec.height;
                setSize(new Dimension(w,y));
                setPreferredSize(new Dimension(w,y));		//2007-05-15 INAS
            }


            //
            //
            //
            public void setPixel(Rectangle d){
                default_rec = d;
            }

            //
            //
            //
            public void paint(Graphics g){
                Dimension d = getSize(null);
//@@                CZSystem.log("CZCMSCCDBMP","WavePanel XView X width[" + d.width + "] height[" + d.height + "]");

                g.setColor(java.awt.Color.black);
                g.fillRect(0,0,d.width,d.height);

                g.setColor(java.awt.Color.gray);

                for(int i = 0 ; i < d.width ; i +=inc){
                    g.drawLine(i,0,i,d.width);

                    String s = new String(i + "");
                    g.drawString(s,i+5,15);
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

//@@            CZSystem.log("CZCMSCCDBMP","UpdateThread START");

            CZSystemQueue   que = new CZSystemQueue(5);
            CZEventAdapter  adp = new CZEventAdapter(que);
            CZEventSender.addCZEventListener(adp);

            while(true){
                try{
                    CZEventCL event = (CZEventCL)que.waitObject();

                    switch(event.getEvent()){
                        case CZEventCL.EV_1023    :
                            //setDefault();
                            get_button.setBackground(CZSystemDefine.BUTTON_WAIT_COL);
                            break;
                         case CZEventCL.EV_8023    :
                            setDefault();
                            get_button.setBackground(CZSystemDefine.BUTTON_NORMAL_COL);
                            break;

                        default         :     break;
                    } // switch end

                }
                catch(Exception e){

                }
            } // while end
        }
    }
}
