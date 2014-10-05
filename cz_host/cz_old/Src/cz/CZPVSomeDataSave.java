package cz;

import java.awt.Component;
import java.awt.Cursor;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.PrintWriter;
import java.io.Serializable;

import java.util.Locale;
import java.util.Properties;
import java.util.Vector;

import javax.swing.DefaultCellEditor;
import javax.swing.JButton;
import javax.swing.ButtonGroup;
import javax.swing.JRadioButton;
import javax.swing.JCheckBox;
import javax.swing.JComboBox;
import javax.swing.JDialog;
import javax.swing.JLabel;
import javax.swing.JOptionPane;
import javax.swing.JPanel;
import javax.swing.JScrollPane;
import javax.swing.JTable;
import javax.swing.JTextField;
import javax.swing.JViewport;
import javax.swing.JFileChooser;
import javax.swing.ListSelectionModel;

import javax.swing.event.ListSelectionEvent;

import javax.swing.plaf.metal.MetalBorders.Flush3DBorder;

import javax.swing.table.AbstractTableModel;
import javax.swing.table.DefaultTableCellRenderer;
import javax.swing.table.DefaultTableColumnModel;
import javax.swing.table.JTableHeader;
import javax.swing.table.TableColumn;

import javax.swing.text.AttributeSet;
import javax.swing.text.BadLocationException;
import javax.swing.text.Document;
import javax.swing.text.PlainDocument;

/*******************************************************************************
 *  複数ＰＶデータ保存
 * @author  (KPK Co.,Ltd.)
 * @version 1.0 (2003/04/01)
 * 
********************************************************************************/
public class CZPVSomeDataSave extends JDialog {
    int     iIdo = 44+4;

    private static CZPVSomeDataSave pvsomedatasave = null;

    private JScrollPane bt_start_scpanel = null;

    private JPanel      bt_panel        = null;
    private BtTable     btTable         = null;
    private JScrollPane bt_scpanel      = null;

    private JPanel      bt_panel2       = null;
    private BtTable2    btTable2        = null;
    private JScrollPane bt_scpanel2     = null;

    private JLabel      lb_sel_proc     = null;
    private JRadioButton rb_Proc[]      = new JRadioButton[10];
    private int         ProcID          = 0;
    private int         SelProcFlg      = 0;

    private JPanel      pv_panel        = null;
    private PVTable     pv_table        = null;

    private JButton     save_button     = null;

    private String      prop_Unit       = null;
    private String      prop_Start      = null;
    private String      prop_End        = null;
    private String      prop_Interval   = null;

    private String      prop_PVItemNo[];

    private JButton     set_save_button = null;
    private JButton     set_read_button = null;

    private Vector      start_list      = null;

    private JComboBox   unit_cbox       = null;

    private NumeralText start_text      = null;
    private NumeralText end_text        = null;
    private NumeralText interval_text   = null;

    private JLabel      start_lab       = null;
    private JLabel      end_lab         = null;
    private JLabel      interval_lab    = null;

    private FileText    file_text       = null;

    private String      save_dir        = null;

    private File        file_           = new File(CZSystem.FILE_SRC_PATH);

    //
    // コンストラクタ
    //
    CZPVSomeDataSave(){
        super();

        setTitle("複数ＰＶデータ保存");
        setSize(890,705+iIdo);
        setResizable(false);
        setModal(true);

        getContentPane().setLayout(null);
        // 他基地参照機能    @20131021
        if(CZSystemDefine.REFERENCE_RUN == CZSystem.getRunLevel()){
            getContentPane().setBackground(CZSystemDefine.DEFAULT_REFERENCE_BACKGROUND_COL);
        }else{
            getContentPane().setBackground(CZSystemDefine.DEFAULT_BACKGROUND_COL);
        }

        try{
            Properties prop =  new Properties();
            FileInputStream pros = new FileInputStream(CZSystemDefine.PROPERTY_FILE);
            prop.load(pros);

            save_dir = prop.getProperty("PV_SAVE_FILE_DIR");
            if(null == save_dir) CZSystem.exit(-1,"CZPVSomeDataSave NO Propertie File null");
            if(1 > save_dir.length()) CZSystem.exit(-1,"CZPVSomeDataSave NO Propertie File name");
        }
        catch( Exception e){
            CZSystem.exit(-1,"CZPVSomeDataSave NO Propertie File");
        }

        try{
            // ----- Property_Fileより 選択PV項目、表示設定を取得する。 --------
            Properties prop =  new Properties();
            FileInputStream pros = new FileInputStream("PVITEM.TXT");
            prop.load(pros);

            // PV項目の設定
            prop_PVItemNo  = new String[128];
            for(int i=0; i < 128 ; i++){
                try {
                    prop_PVItemNo[i]   = prop.getProperty("PVDATA" + (i+1));
                } catch (Exception e) {
                    prop_PVItemNo[i]   = new String("");
                }
            }
            // 表示設定
            prop_Unit     = prop.getProperty("SELUNIT");
            prop_Start    = prop.getProperty("START");
            prop_End      = prop.getProperty("END");
            prop_Interval = prop.getProperty("INTERVAL");

        } catch( Exception e ) {
            //プロパティ取得でエラーの時は、終了する。
            CZSystem.exit(-1,"CZPVSomeDataSave NO Propertie File");
        }

		/* 上段（左）のパネル */
		bt_panel = new JPanel();
		bt_panel.setLayout(null);
		bt_panel.setBorder(new Flush3DBorder());
		bt_panel.setBackground(java.awt.Color.lightGray);
		bt_panel.setBounds(20, 20, 402, 360);
		getContentPane().add(bt_panel);

		bt_scpanel = new JScrollPane();
		bt_scpanel.setBounds(0, 0, 402, 360);
		bt_panel.add(bt_scpanel);

		/* 上段（右）のパネル */
		bt_panel2 = new JPanel();
		bt_panel2.setLayout(null);
		bt_panel2.setBorder(new Flush3DBorder());
		bt_panel2.setBackground(java.awt.Color.lightGray);
		bt_panel2.setBounds(445, 20, 420, 360);
		getContentPane().add(bt_panel2);

		bt_scpanel2 = new JScrollPane();
		bt_scpanel2.setBounds(0, 0, 420 ,360);
		bt_panel2.add(bt_scpanel2);

		lb_sel_proc = new JLabel("プロセス選択");
		lb_sel_proc.setBounds(20, 390, 120, 24);
		lb_sel_proc.setLocale(new Locale("ja","JP"));
		lb_sel_proc.setFont(new java.awt.Font("dialog", 1, 18));
		lb_sel_proc.setForeground(java.awt.Color.black);
		getContentPane().add(lb_sel_proc);

		ButtonGroup rb_proc_grp = new ButtonGroup();

		/* プロセス選択ラジオボタン */
		for(int i = 0; i < 10; i++){
			String proc = CZSystem.getProcName2(i);
			rb_Proc[i] = new JRadioButton(proc);
			rb_Proc[i].setBounds(20+(i*85), 420, 80, 24);
			rb_Proc[i].setLocale(new Locale("ja","JP"));
			rb_Proc[i].setFont(new java.awt.Font("dialog", 1, 16));
			rb_Proc[i].setBorder(new Flush3DBorder());
			rb_Proc[i].setForeground(java.awt.Color.black);
			rb_Proc[i].addActionListener(new SelProc());
	        // 他基地参照機能    @20131021
	        if(CZSystemDefine.REFERENCE_RUN == CZSystem.getRunLevel()){
	            rb_Proc[i].setBackground(CZSystemDefine.DEFAULT_REFERENCE_BACKGROUND_COL);
	        }else{
	            rb_Proc[i].setBackground(CZSystemDefine.DEFAULT_BACKGROUND_COL);
	        }
			getContentPane().add(rb_Proc[i]);
			rb_proc_grp.add(rb_Proc[i]);
		}

        /* 下段のパネル */
        pv_panel = new JPanel();
        pv_panel.setLayout(null);
        pv_panel.setBounds(20, 400+iIdo, 845, 230);
        pv_panel.setBorder(new Flush3DBorder());
        pv_panel.setBackground(java.awt.Color.lightGray);
        getContentPane().add(pv_panel);

        pv_table = new PVTable();
        JScrollPane panel = new JScrollPane(pv_table);
        panel.setBounds(10, 20, 488 ,187);
        pv_panel.add(panel);

        unit_cbox = new JComboBox();
        unit_cbox.setBounds(510, 20, 80, 24);
        unit_cbox.setLocale(new Locale("ja","JP"));
        unit_cbox.setFont(new java.awt.Font("dialog", 0, 12));
        unit_cbox.setForeground(java.awt.Color.black);
        unit_cbox.addItem("時間");
        unit_cbox.addItem("長さ");
        unit_cbox.setFocusable(false);
        unit_cbox.addActionListener(new ChgUnit());
        pv_panel.add(unit_cbox);

        JLabel  lab = new JLabel("開始",JLabel.CENTER);
        lab.setBounds(590, 20, 50, 24);
        lab.setLocale(new Locale("ja","JP"));
        lab.setFont(new java.awt.Font("dialog", 0, 12));
        lab.setBorder(new Flush3DBorder());
        lab.setForeground(java.awt.Color.black);
        pv_panel.add(lab);

        lab = new JLabel("終了",JLabel.CENTER);
        lab.setBounds(590, 50, 50, 24);
        lab.setLocale(new Locale("ja","JP"));
        lab.setFont(new java.awt.Font("dialog", 0, 12));
        lab.setBorder(new Flush3DBorder());
        lab.setForeground(java.awt.Color.black);
        pv_panel.add(lab);

        lab = new JLabel("間隔",JLabel.CENTER);
        lab.setBounds(590, 80, 50, 24);
        lab.setLocale(new Locale("ja","JP"));
        lab.setFont(new java.awt.Font("dialog", 0, 12));
        lab.setBorder(new Flush3DBorder());
        lab.setForeground(java.awt.Color.black);
        pv_panel.add(lab);

        start_text = new NumeralText(6,false);
        start_text.setBounds(640, 20, 80, 24);
        start_text.setText("0");
        pv_panel.add(start_text);

        end_text = new NumeralText(6,false);
        end_text.setBounds(640, 50, 80, 24);
        end_text.setText("0");
        pv_panel.add(end_text);

        interval_text = new NumeralText(6,false);
        interval_text.setBounds(640, 80, 80, 24);
        interval_text.setText("0");
        pv_panel.add(interval_text);

        start_lab = new JLabel("X 10s",JLabel.CENTER);
        start_lab.setBounds(720, 20, 60, 24);
        start_lab.setLocale(new Locale("ja","JP"));
        start_lab.setFont(new java.awt.Font("dialog", 0, 12));
        start_lab.setBorder(new Flush3DBorder());
        start_lab.setForeground(java.awt.Color.black);
        pv_panel.add(start_lab);

        end_lab = new JLabel("X 10s",JLabel.CENTER);
        end_lab.setBounds(720, 50, 60, 24);
        end_lab.setLocale(new Locale("ja","JP"));
        end_lab.setFont(new java.awt.Font("dialog", 0, 12));
        end_lab.setBorder(new Flush3DBorder());
        end_lab.setForeground(java.awt.Color.black);
        pv_panel.add(end_lab);

        interval_lab = new JLabel("X 10s",JLabel.CENTER);
        interval_lab.setBounds(720, 80, 60, 24);
        interval_lab.setLocale(new Locale("ja","JP"));
        interval_lab.setFont(new java.awt.Font("dialog", 0, 12));
        interval_lab.setBorder(new Flush3DBorder());
        interval_lab.setForeground(java.awt.Color.black);
        pv_panel.add(interval_lab);

        JTextField hed = new JTextField(save_dir);
        hed.setBounds(510, 125, 270, 24);
        hed.setEnabled(false);
        pv_panel.add(hed);

        file_text = new FileText();
        file_text.setBounds(510, 150, 270, 24);
        pv_panel.add(file_text);

        save_button = new JButton("実績保存");
        save_button.setBounds(510, 183, 100, 24);
        save_button.setLocale(new Locale("ja","JP"));
        save_button.setFont(new java.awt.Font("dialog", 0, 18));
        save_button.setBorder(new Flush3DBorder());
        save_button.setForeground(java.awt.Color.black);
        save_button.addActionListener(new SaveButton());
        pv_panel.add(save_button);

        // ======================================== [設定読込]ボタン ==================================
        set_read_button = new JButton("設定読込");
        set_read_button.setBounds(620, 183, 100, 24);
        set_read_button.setLocale(new Locale("ja","JP"));
        set_read_button.setFont(new java.awt.Font("dialog", 0, 18));
        set_read_button.setBorder(new Flush3DBorder());
        set_read_button.setForeground(java.awt.Color.black);
        set_read_button.addActionListener(
          new ActionListener() {
              public void actionPerformed(ActionEvent evt) {
                  JFileChooser chooser = new JFileChooser(file_);
                  int ret = chooser.showOpenDialog(pvsomedatasave);
                  if ( ret == JFileChooser.APPROVE_OPTION ) {
                      file_ = chooser.getSelectedFile();        // ファイル名を取得する
                      Properties prop = new Properties();       // プロパティを生成する
                      try {
                          FileInputStream in = new FileInputStream(file_);
                          prop.load( in );                      //プロパティを取得する。
                          in.close();
                          prop.list(System.out);

                          // 表示設定の読込
                          prop_Unit     = prop.getProperty("SELUNIT");
                          prop_Start    = prop.getProperty("START");
                          prop_End      = prop.getProperty("END");
                          prop_Interval = prop.getProperty("INTERVAL");

                          // PV項目の読込
                          for(int i=0; i < 128 ; i++){
                              try {
                                  prop_PVItemNo[i]   = prop.getProperty("PVDATA" + (i+1));
                              } catch (Exception e) {
                                  prop_PVItemNo[i]   = new String("");
                              }
                          }

                          pv_table.clearSelection(); // 全PV項目選択クリア
                          // ファイルから読み込んだPV項目を選択
                          for(int i = 0; i < 128 ; i++){
                              if(prop_PVItemNo[i].equals("1")){
                                  CZSystem.log("CZPVSomeDataSave","PVDATA" + i + " : " + prop_PVItemNo[i]);
                                  pv_table.addRowSelectionInterval(i,i);
                              }
                          }

                          // 時間 or 長さ 設定
                          if(prop_Unit.equals("1")){
                              unit_cbox.setSelectedIndex(1);
                          }else{
                              unit_cbox.setSelectedIndex(0);
                          }

                          // 開始・終了・間隔 設定
                          start_text.setText(prop_Start);
                          end_text.setText(prop_End);
                          interval_text.setText(prop_Interval);

                      } catch ( IOException ex ) {
                          CZSystem.log("CZPVSomeDataSave","PVITEM Property Fileがロードできませんでした。");
                          return;
                      }
                  }
              }
          }
        );
        pv_panel.add(set_read_button);

        // ======================================== [設定保存]ボタン ==================================
        set_save_button = new JButton("設定保存");
        set_save_button.setBounds(730, 183, 100, 24);
        set_save_button.setLocale(new Locale("ja","JP"));
        set_save_button.setFont(new java.awt.Font("dialog", 0, 18));
        set_save_button.setBorder(new Flush3DBorder());
        set_save_button.setForeground(java.awt.Color.black);
        set_save_button.addActionListener(
          new ActionListener() {
              public void actionPerformed(ActionEvent evt)
              {
                  JFileChooser chooser = new JFileChooser(file_);
                  int ret = chooser.showSaveDialog(pvsomedatasave);
                  if (ret == JFileChooser.APPROVE_OPTION) {
                      file_ = chooser.getSelectedFile();            // ファイル名を取得する
                      Properties prop = new Properties();           // プロパティを生成する
                      // 表示の設定
                      prop.setProperty(new String("SELUNIT"),  new String("" + unit_cbox.getSelectedIndex()) );
                      prop.setProperty(new String("START"),    new String("" + start_text.getText()));
                      prop.setProperty(new String("END"),      new String("" + end_text.getText())  );
                      prop.setProperty(new String("INTERVAL"), new String("" + interval_text.getText()) );

                      // PV項目の設定
                      for (int i = 0; i < 128; i++) {
                        if(pv_table.isRowSelected(i) == true){
                            prop.setProperty(new String("PVDATA" + (i+1)),  new String("1"));
                        }else{
                            prop.setProperty(new String("PVDATA" + (i+1)),  new String("0"));
                        }
                      }
                      //---------- ファイルに保存する  ----------
                      try {
                          FileOutputStream out = new FileOutputStream(file_);
                          prop.store(out, "");
                          out.flush();
                          out.close();
                      } catch (IOException ex) {
                          JOptionPane.showMessageDialog(
                            pvsomedatasave,
                            new String("保存できませんでした。"),
                            new String("保存"),
                            JOptionPane.WARNING_MESSAGE);
                          return;
                      }
                      JOptionPane.showMessageDialog(
                        pvsomedatasave,
                        new String("保存しました。"),
                        new String("保存"),
                        JOptionPane.INFORMATION_MESSAGE);
                      return;
                  }
              }
          }
        );
        pv_panel.add(set_save_button);


    }

    //
    // 画面初期表示
    public boolean setDefault(){
        // 保存するFile名をクリアする。
        clearFileName();

//        rb_Proc[0].setSelected(true);
        CZSystem.log("CZPVSomeDataSave","ラジオボタンVAC有効");

        btTable = new BtTable();

        JTableHeader tabHead = btTable.getTableHeader();
        tabHead.setReorderingAllowed(false);
        bt_scpanel.setViewportView(btTable);

        btTable2 = new BtTable2();

        JTableHeader tabHead2 = btTable2.getTableHeader();
        tabHead2.setReorderingAllowed(false);
        bt_scpanel2.setViewportView(btTable2);

        clearBtSelect();
        pv_table.clearSelection(); // 全PV項目選択クリア

        return true;
    }

    //
    // カーソルを設定する。
    private void setCur(Cursor cu){
        setCursor(cu);
    }

    //
    // カーソルを取得する。
    private Cursor getCur(){
        return getCursor();
    }

    //
    // File名をクリアする。
    public void clearFileName(){
        String text = "";
        file_text.setText(text);
        return;
    }

    //
    // 選択バッチのクリア
    public void clearBtSelect(){
        for(int i = 0 ; i < btTable.getRowCount() ; i++){
            Boolean bFlg = new Boolean(false);
            btTable.setValueAt(bFlg,i,0);
        }
            btTable.repaint();

        for(int i = 0 ; i < btTable2.getRowCount() ; i++){
            Boolean bFlg = new Boolean(false);
            btTable2.setValueAt(bFlg,i,0);
        }
            btTable2.repaint();

    }

    //
    // メッセージの表示
    //
    private boolean infoMsg(Object msg[]){
        JOptionPane.showMessageDialog(null,msg,
                                    "複数PVデータ保存",
        JOptionPane.INFORMATION_MESSAGE);
        return true;
    }


    //
    // PVデータを保存する。
    public boolean savePVData(int proc, String bt, String spec, int ichg, int rcp){

        CZSystem.log("CZPVSomeDataSave savePVData()"," [" + bt + "]");
        CZSystem.log("CZPVSomeDataSave savePVData()"," [" + spec + "]");
        CZSystem.log("CZPVSomeDataSave savePVData()"," [" + ichg + "]");
        CZSystem.log("CZPVSomeDataSave savePVData()"," [" + rcp + "]");

        String db = CZSystem.getDBName();

        Vector procList = CZSystem.getPvProcNo(db,proc,bt,spec,ichg,rcp);
        if(procList == null){
            return false;
        }else{
            CZSystem.log("CZPVSomeDataSave savePVData() procList size "," [" + procList.size() + "]");
        }

        for(int row = 0; row < procList.size(); row++){
            CZSaveDataProcList plist   = (CZSaveDataProcList)procList.elementAt(row);
            CZSystem.log("CZPVSomeDataSave savePVData() plist.p_no "," [" + plist.p_no + "]");
            CZSystem.log("CZPVSomeDataSave savePVData() plist.sp_no "," [" + plist.sp_no + "]");
            CZSystem.log("CZPVSomeDataSave savePVData() plist.p_renban "," [" + plist.p_renban + "]");

            String file_name = bt.trim() + "_Proc-" + plist.p_no + "_P連番-" + plist.p_renban + "_品種-" + spec.trim() + "_初仕込-" + ichg + "_T2-" + rcp + ".csv";
            if(1 > file_name.length()) return false;

            // PVデータを保持する領域を確保し、クリアする。
            boolean data_no[] = new boolean[CZSystemDefine.PV_MAX_LENGTH];
            for(int i = 0 ; i < CZSystemDefine.PV_MAX_LENGTH ; i++) data_no[i] = false;

            int list[];
            int no;
            list = pv_table.getSelectedRows();
            if(1 > list.length) return false;

            for(int i = 0 ; i < list.length ; i++){
                CZSystem.log("CZPVSomeDataSave savePVSomeData()","savePVSomeData List[" + i + "][" + list[i] + "]");

                no = list[i];
                if(no < CZSystemDefine.PV_MAX_LENGTH) data_no[no] = true;
            }

            String view = CZSystem.getViewName(db,bt);

            Vector pv_data = CZSystem.getPVSomeData(db,view,plist.p_no,plist.p_renban,data_no);
            if(null == pv_data) return false;
            if(1 > pv_data.size()) return false;
            String jhed = createJHed(data_no);
            String hed = createHed(data_no);
//@@        CZSystem.log("CZPVSomeDataSave savePVData()","Head [" + hed + "]" );

            writeFile(file_name,jhed,hed,pv_data,data_no);
        }

        return true;
    }
/*
    //
    // PVデータを保存する。
    public boolean savePVData(int row){
        // File名が設定されていない時はエラーにする。
        String file_name = file_text.getText();
        if(1 > file_name.length()) return false;
        // バッチ開始情報が無い時はエラーにする。
        if(null == start_list) return false;
        if(0 > row) return false;

//@@        CZSystem.log("CZPVSomeDataSave savePVData()"," [" + row + "]");
        // 有効な開始情報を選択していない時はエラーにする。
        CZSystemStart st = (CZSystemStart)start_list.elementAt(row);
        if(null == st) return false;

//@@        CZSystem.log("CZPVSomeDataSave savePVData()",
//@@            " [" + st.batch + "][" + st.p_no + "][" + st.p_renban + "][" + st.p_start + "]");
        // PVデータを保持する領域を確保し、クリアする。
        boolean data_no[] = new boolean[CZSystemDefine.PV_MAX_LENGTH];
        for(int i = 0 ; i < CZSystemDefine.PV_MAX_LENGTH ; i++) data_no[i] = false;

        int list[];
        int no;
        list = pv_table.getSelectedRows();
        if(1 > list.length) return false;

        for(int i = 0 ; i < list.length ; i++){
            CZSystem.log("CZPVSomeDataSave savePVData()","savePVData List[" + i + "][" + list[i] + "]");

            no = list[i];
            if(no < CZSystemDefine.PV_MAX_LENGTH) data_no[no] = true;
        }

        String db = CZSystem.getDBName();
        String view = CZSystem.getViewName(db,st.batch);

        Vector pv_data = CZSystem.getPVData(db,view,st.p_renban,data_no);
        if(null == pv_data) return false;
        if(1 > pv_data.size()) return false;
        String jhed = createJHed(data_no);
        String hed = createHed(data_no);
//@@        CZSystem.log("CZPVSomeDataSave savePVData()","Head [" + hed + "]" );

        writeFile(jhed,hed,pv_data,data_no);

        return true;
    }
*/

    //
    //
    //
    private int writeFile(String fname,String jhed,String hed,Vector pv_data,boolean data_no[]){
        int write_count    = 1;

        File file          = new File(save_dir,fname);
        FileOutputStream s = null;
        PrintWriter pr     = null;

        float pos          = 0;
        float next_pos     = 0;

        int start          = 0;
        int end            = 0;
        int inc            = 0;

        start = start_text.getValue();
        end   = end_text.getValue();
        inc   = interval_text.getValue();

        int index = unit_cbox.getSelectedIndex();
        switch(index){
            case 0:
            default:
                start = start * 10;
                end   = end * 10;
                inc   = inc * 10;
                break;
            case 1:
                break;
        }

        try{

            s  = new FileOutputStream(file);
            pr = new PrintWriter(s);
            // Header部を書く。
            pr.println(jhed);
            // Header部を書く。
            pr.println(hed);
            // PVデータを書く。
            for(int i = 0 ; i < pv_data.size() ; i++){
                CZSystemPVData pv = (CZSystemPVData)pv_data.elementAt(i);

                switch(index){
                case 0:
                default:
                    pos = pv.p_time;
                    break;
                case 1:
                    pos = pv.p_length;  
                    break;
                }

                if(start  > pos) continue;
                if((end != 0) && (end < pos)) break;
                if(1 == write_count) next_pos = start;

                if(next_pos <= pos){
                    writeRec(pr,pv,data_no);
                    write_count++;
                    next_pos += inc;
                }
            }
        }
        catch(IOException e){
            CZSystem.log("CZPVSomeDataSave writeFile()","[" + write_count + "][" + e + "]");
            if(null != pr) pr.close();
            return -1;
        }

        if(null != pr) pr.close();

        write_count--;
        return write_count;
    }

    //
    // PVデータをCSV形式で書く。
    private boolean writeRec(PrintWriter pr,CZSystemPVData pv,boolean data_no[]){

        StringBuffer rec = new StringBuffer(3000);  

        try{
            rec.append(pv.p_no);
            rec.append("," + pv.sp_no);
            rec.append("," + pv.p_renban);
            rec.append("," + pv.p_time);
            rec.append("," + pv.sp_time);
            rec.append("," + pv.p_date);
            rec.append("," + pv.h_ontime);
            rec.append("," + pv.hk_renban);

            for(int j = 0 ; j < data_no.length ; j++){
                if(!data_no[j]) continue;
                rec.append("," + pv.data[j]);
            }

            pr.println(rec);
        }
        catch(Exception e){
            CZSystem.log("CZPVSomeDataSave writeRec()","[" + e + "]");
            return false;
        }
        return true;
    }

    //
    // Header部1を作成する。日本語名称部
    private String createJHed(boolean v[]){

        StringBuffer buf = new StringBuffer(2000);

        buf.append("プロセスNo,サブプロセスNo,プロセス連番,プロセス時間,サブプロセス時間,採取日時,メインヒータ電源オン時間,引上げ条件内連番");
        for(int i = 0 ; i < v.length ; i++){
            if(!v[i]) continue;
            CZSystemPVName name = CZSystem.getPVName(i);
            if(null == name) break;
            buf.append("," + name.j_name.trim());
        }
        String ret = buf.toString();
        return ret;
    }

    //
    // Header部2を作成する。
    private String createHed(boolean v[]){

        StringBuffer buf = new StringBuffer(2000);

        buf.append("p_no,sp_no,p_renban,p_time,sp_time,p_date,h_ontime,hk_renban");
        for(int i = 0 ; i < v.length ; i++){
            if(!v[i]) continue;
            CZSystemPVName name = CZSystem.getPVName(i);
            if(null == name) break;
            buf.append("," + name.k_name.trim());
        }
        String ret = buf.toString();
        return ret;
    }

    /***************************************************************************
     *  時間or長さ選択ComboBoxのリスナー
     *  (選択に合わせて単位表示を変える) X 10s <--> mm  
     ***************************************************************************/
    class ChgUnit implements ActionListener {
        public void actionPerformed(ActionEvent ev){

            int index = unit_cbox.getSelectedIndex();

//@@            CZSystem.log("CZPVSomeDataSave ChgUnit()","[" + index + "]" );

            switch(index){
            case 0:
            default:
                start_lab.setText("X 10s");
                end_lab.setText("X 10s");
                interval_lab.setText("X 10s");
                break;
            case 1:
                start_lab.setText("mm");
                end_lab.setText("mm");
                interval_lab.setText("mm");
                break;
            }
        }
    }

    /***************************************************************************
     *
     *  ラジオボタンのリスナー
     *
     ***************************************************************************/
	class SelProc implements ActionListener {
		public void actionPerformed(ActionEvent e){
		
			for(int rec = 0; rec < 10; rec++){
				if(true == rb_Proc[rec].isSelected()){
					CZSystem.log("CZPVSomeDataSave SelProc","select RadioButton rec [" + rec + "]");
					CZSystem.log("CZPVSomeDataSave SelProc","select RadioButton [" + rb_Proc[rec].getText() + "]");
					ProcID = rec + 1;
					CZSystem.log("CZPVSomeDataSave SelProc","select RadioButton ProcID [" + ProcID + "]");
					SelProcFlg = 1;
					return;
				}
			}
		}
	}

    /***************************************************************************
     *
     *  保存ボタンのリスナー
     *
     ***************************************************************************/
    class SaveButton implements ActionListener {
        public void actionPerformed(ActionEvent ev){

            Object  oDt;
            boolean flg;
            String  Batch;
            String  Hinshu;
            int     I_sikomi;
            int     No_hikiage;

            if(SelProcFlg == 0){
                Object msg[] = {"複数PVデータ保存",
                                "プロセスを選択して下さい"};
                infoMsg(msg);
                return;
            }

            for(int i = 0; i < btTable.getRowCount(); i++){
                oDt = btTable.getValueAt(i,0);
                flg = ((Boolean)oDt).booleanValue();
                if(flg == true){
                    Batch      = btTable.getValueAt(i,2).toString();
                    Hinshu     = btTable.getValueAt(i,3).toString();
                    String s   = btTable.getValueAt(i,4).toString();
                    I_sikomi   = Integer.valueOf(s).intValue();
                    String n   = btTable.getValueAt(i,5).toString();
                    No_hikiage = Integer.valueOf(n).intValue();

                    CZSystem.log("CZPVSomeDataSave SaveButton","valueChanged [" + ProcID + "]");
                    CZSystem.log("CZPVSomeDataSave SaveButton","valueChanged [" + Batch + "]");
                    CZSystem.log("CZPVSomeDataSave SaveButton","valueChanged [" + Hinshu + "]");
                    CZSystem.log("CZPVSomeDataSave SaveButton","valueChanged [" + I_sikomi + "]");
                    CZSystem.log("CZPVSomeDataSave SaveButton","valueChanged [" + No_hikiage + "]");

                    savePVData(ProcID,Batch,Hinshu,I_sikomi,No_hikiage);
                }
            }

            for(int i = 0; i < btTable2.getRowCount(); i++){
                oDt = btTable2.getValueAt(i,0);
                flg = ((Boolean)oDt).booleanValue();
                if(flg == true){
                    Batch      = btTable2.getValueAt(i,2).toString();
                    Hinshu     = btTable2.getValueAt(i,3).toString();
                    String s   = btTable2.getValueAt(i,4).toString();
                    I_sikomi   = Integer.valueOf(s).intValue();
                    String n   = btTable2.getValueAt(i,5).toString();
                    No_hikiage = Integer.valueOf(n).intValue();

                    CZSystem.log("CZPVSomeDataSave SaveButton","valueChanged [" + ProcID + "]");
                    CZSystem.log("CZPVSomeDataSave SaveButton","valueChanged [" + Batch + "]");
                    CZSystem.log("CZPVSomeDataSave SaveButton","valueChanged [" + Hinshu + "]");
                    CZSystem.log("CZPVSomeDataSave SaveButton","valueChanged [" + I_sikomi + "]");
                    CZSystem.log("CZPVSomeDataSave SaveButton","valueChanged [" + No_hikiage + "]");

                    savePVData(ProcID,Batch,Hinshu,I_sikomi,No_hikiage);
                }
            }

/*
            v = bt_scpanel.getViewport();
            t = (JTable)v.getView();
            if(null == t) return;
            int bt_row = t.getSelectedRow();

            v = bt_start_scpanel.getViewport();
            t = (JTable)v.getView();
            if(null == t) return;
            int bt_start_row = t.getSelectedRow();
*/
//@@            CZSystem.log("CZPVSomeDataSave SaveButton","valueChanged [" + bt_row + "][" + bt_start_row + "]");

            // カーソルをWaitに設定する。
            Cursor cu_tmp = getCur();
            Cursor cu     = new Cursor(Cursor.WAIT_CURSOR);
            setCur(cu);
            // PVデータを保存する。
            //savePVData(bt_start_row);
            // カーソルを元に戻す。
            setCur(cu_tmp);

        }
    }

    /***************************************************************************
     *
     *       ＰＶ引上げバッチ一覧（左パネル）
     *
     ***************************************************************************/
    class BtTable extends JTable {

        private Vector   bt_list = null;

        private BtTblMdl model   = null;

        BtTable(){
            super();

            try{
                setName("BtTable");
                setAutoCreateColumnsFromModel(true);
                setSelectionMode(ListSelectionModel.SINGLE_SELECTION);

                setLocale(new Locale("ja","JP"));
                setFont(new java.awt.Font("dialog", 0, 12));
                setRowHeight(17);

                bt_list  = CZSystem.getPVDataBtList(CZSystem.getDBName());

                model = new BtTblMdl(bt_list);
                setModel(model);

                DefaultTableColumnModel cmdl = (DefaultTableColumnModel)getColumnModel();
                TableColumn colum = null;

                // Umu
                colum = cmdl.getColumn(0);
                colum.setMaxWidth(20);
                colum.setMinWidth(20);
                colum.setWidth(20);
                JCheckBox check = new JCheckBox();
                DefaultCellEditor editor = new DefaultCellEditor(check);
                colum.setCellEditor( editor );

                // No
                colum = cmdl.getColumn(1);
                colum.setMaxWidth(40);
                colum.setMinWidth(40);
                colum.setWidth(40);

                // Batch
                colum = cmdl.getColumn(2);
                colum.setMaxWidth(140);
                colum.setMinWidth(140);
                colum.setWidth(140);

                // 品種
                colum = cmdl.getColumn(3);
                colum.setMaxWidth(80);
                colum.setMinWidth(80);
                colum.setWidth(80);

                // 仕込量
                colum = cmdl.getColumn(4);
                colum.setMaxWidth(80);
                colum.setMinWidth(80);
                colum.setWidth(80);

                // T2
                colum = cmdl.getColumn(5);
                colum.setMaxWidth(40);
                colum.setMinWidth(40);
                colum.setWidth(40);

            }
            catch (Throwable e) {
                
            }
        }

        //
        //
        //
        public void valueChanged(ListSelectionEvent e){
            super.valueChanged(e);

            if(e.getValueIsAdjusting()) return;
            int row = getSelectedRow();
//            int col = getSelectedCol();

			if(0 <= row){
				CZSystem.log("CZPVSomeDataSave BtTable","valueChanged [" + row + "]");
/*				CZSystem.log(CZCalcSystem.INFO, "row = " + row  );
				CZCalcSystem.log(CZCalcSystem.INFO, "DATA = " + getValueAt(row,0));
				CZSystemRoBt bt = (CZSystemRoBt)bt_list.elementAt(row);
				CZCalcSystem.log(CZCalcSystem.INFO, "UMU = " + bt.umuFlg);*/
			}
		}
        //
        //
        //
        public void setData(){
//        CZCalcSystem.log( CZCalcSystem.INFO, "CZPVSomeDataSave setData()[" + gr + "][" + tbl + "]");
        }

    }

    /***************************************************************************
     *
     *       ＰＶ引上げバッチ一覧：モデル（左パネル）
     *
     ***************************************************************************/
    public class BtTblMdl extends AbstractTableModel {

        private int TBL_ROW     = 0;
        final   int TBL_COL     = 6;
        private Vector  bt_list = null;

        final String[] names = {" " , " # " , "Bt" , "品種" , "仕込量" , "T2" };

        private Object  data[][];

        BtTblMdl(Vector v){
            super();

            bt_list = v;
            TBL_ROW = bt_list.size();

            data = new Object[TBL_ROW][TBL_COL];

            for(int i = 0 ; i < TBL_ROW ; i++){
                CZPVDataBtList bt = (CZPVDataBtList)bt_list.elementAt(i);
                if(null == bt) break;

                if(bt.flg == 0){
                    data[i][0] = new Boolean(false);
                }else{
                    data[i][0] = new Boolean(true);
                }
                data[i][1] = new Integer(i+1);
                data[i][2] = bt.batch;
                data[i][3] = bt.hinshu;
                data[i][4] = bt.i_sikomi;
                data[i][5] = bt.no_hikiage;
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
            if( col == 0 ){
                return true;
            }
            return false;
        }

        public void setValueAt(Object aValue, int row, int column){
            data[row][column] = aValue;
        }
    }


    /***************************************************************************
     *
     *       ＰＶ引上げバッチ一覧（右パネル）
     *
     ***************************************************************************/
    class BtTable2 extends JTable {

        private Vector   bt_list2 = null;

        private BtTblMdl2 model   = null;

        BtTable2(){
            super();

            try{
                setName("BtTable2");
                setAutoCreateColumnsFromModel(true);
                setSelectionMode(ListSelectionModel.SINGLE_SELECTION);

                setLocale(new Locale("ja","JP"));
                setFont(new java.awt.Font("dialog", 0, 12));
                setRowHeight(17);

                bt_list2  = CZSystem.getPVDataBtList2(CZSystem.getDBName());

                model = new BtTblMdl2(bt_list2);
                setModel(model);

                DefaultTableColumnModel cmdl = (DefaultTableColumnModel)getColumnModel();
                TableColumn colum = null;

                // Umu
                colum = cmdl.getColumn(0);
                colum.setMaxWidth(20);
                colum.setMinWidth(20);
                colum.setWidth(20);
                JCheckBox check = new JCheckBox();
                DefaultCellEditor editor = new DefaultCellEditor(check);
                colum.setCellEditor( editor );

                // No
                colum = cmdl.getColumn(1);
                colum.setMaxWidth(40);
                colum.setMinWidth(40);
                colum.setWidth(40);

                // Batch
                colum = cmdl.getColumn(2);
                colum.setMaxWidth(140);
                colum.setMinWidth(140);
                colum.setWidth(140);

                // 品種
                colum = cmdl.getColumn(3);
                colum.setMaxWidth(80);
                colum.setMinWidth(80);
                colum.setWidth(80);

                // 仕込量
                colum = cmdl.getColumn(4);
                colum.setMaxWidth(80);
                colum.setMinWidth(80);
                colum.setWidth(80);

                // T2
                colum = cmdl.getColumn(5);
                colum.setMaxWidth(40);
                colum.setMinWidth(40);
                colum.setWidth(40);

            }
            catch (Throwable e) {
                
            }
        }

        //
        //
        //
        public void valueChanged(ListSelectionEvent e){
            super.valueChanged(e);

            if(e.getValueIsAdjusting()) return;
            int row = getSelectedRow();
//            int col = getSelectedCol();
/*
			if(0 <= row){
				CZCalcSystem.log(CZCalcSystem.INFO, "row = " + row  );
				CZCalcSystem.log(CZCalcSystem.INFO, "DATA = " + getValueAt(row,0));
				CZSystemRoBt bt = (CZSystemRoBt)bt_list.elementAt(row);
				CZCalcSystem.log(CZCalcSystem.INFO, "UMU = " + bt.umuFlg);*/
			}
        //
        //
        //
        public void setData(){
//            CZCalcSystem.log( CZCalcSystem.INFO, "CZPVSomeDataSave setData()[" + gr + "][" + tbl + "]");

        }

    }

    /***************************************************************************
     *
     *       ＰＶ引上げバッチ一覧：モデル（右パネル）
     *
     ***************************************************************************/
    public class BtTblMdl2 extends AbstractTableModel {

        private int TBL_ROW      = 0;
        final   int TBL_COL      = 6;
        private Vector  bt_list2 = null;

        final String[] names = {" " , " # " , "Bt" , "品種" , "仕込量" , "T2" };

        private Object  data[][];

        BtTblMdl2(Vector v){
            super();

            bt_list2 = v;
            TBL_ROW = bt_list2.size();

            data = new Object[TBL_ROW][TBL_COL];

            for(int i = 0 ; i < TBL_ROW ; i++){
                CZPVDataBtList bt = (CZPVDataBtList)bt_list2.elementAt(i);
                if(null == bt) break;

                if(bt.flg == 0){
                    data[i][0] = new Boolean(false);
                }else{
                    data[i][0] = new Boolean(true);
                }
                data[i][1] = new Integer(i+21);
                data[i][2] = bt.batch;
                data[i][3] = bt.hinshu;
                data[i][4] = bt.i_sikomi;
                data[i][5] = bt.no_hikiage;
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
            if( col == 0 ){
                return true;
            }
            return false;
        }

        public void setValueAt(Object aValue, int row, int column){
            data[row][column] = aValue;
        }
    }

    /***************************************************************************
     *
     *       ＰＶ名一覧
     *
     ***************************************************************************/
    class PVTable extends JTable {

        private PVTblMdl model = null;

        PVTable(){
            super();

            try{
                setName("PVTable");
                setAutoCreateColumnsFromModel(true);
                setSelectionMode(ListSelectionModel.MULTIPLE_INTERVAL_SELECTION);

                setLocale(new Locale("ja","JP"));
                setFont(new java.awt.Font("dialog", 0, 12));
                setRowHeight(17);

                model = new PVTblMdl();
                setModel(model);
                DefaultTableColumnModel cmdl = (DefaultTableColumnModel)getColumnModel();
                TableColumn colum = null;

                for(int i=0; i < 128 ; i++){
                    if(prop_PVItemNo[i].equals("1")){
                        addRowSelectionInterval(i,i);
                    }
                }

                // No
                colum = cmdl.getColumn(0);
                colum.setMaxWidth(40);
                colum.setMinWidth(40);
                colum.setWidth(40);

                // 項目名
                colum = cmdl.getColumn(1);
                colum.setMaxWidth(80);
                colum.setMinWidth(80);
                colum.setWidth(80);

                // 単位
                colum = cmdl.getColumn(2);
                colum.setMaxWidth(70);
                colum.setMinWidth(70);
                colum.setWidth(70);

                // 説明
                colum = cmdl.getColumn(3);
                colum.setMaxWidth(280);
                colum.setMinWidth(280);
                colum.setWidth(280);

            }
            catch (Throwable e) {
                CZSystem.handleException(e);
            }
        }

        //
        //
        public void valueChanged(ListSelectionEvent e){
            //CZSystem.log("CZPVSomeDataSave","PVTable");
            super.valueChanged(e);
        }

        //
        //
        public void setData(int gr,int tbl){

//            CZSystem.log("CZPVSomeDataSave","PVTable setData() [" + gr + "][" + tbl + "]");
        }
    }

    /***************************************************************************
     *
     *       ＰＶ名一覧：モデル
     *
     ***************************************************************************/
    public class PVTblMdl extends AbstractTableModel {

        private int TBL_ROW     = CZSystemDefine.PV_MAX_LENGTH;
        final   int TBL_COL     = 4;

        final String[] names = {" # " , "項目名" , "単位" , "説明" };

        private Object  data[][];

        PVTblMdl(){
            super();

            data = new Object[TBL_ROW][TBL_COL];

            String empty   = new String("");

            for(int i = 0 ; i < TBL_ROW ; i++){
                CZSystemPVName name = CZSystem.getPVName(i);
                if(null == name) break;

                data[i][0] = new Integer(name.k_no);
                data[i][1] = name.k_name;
                data[i][2] = name.unit;
                data[i][3] = name.j_name;
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
     *       数値を設定するTextField
     *
     */
    public class NumeralText extends JTextField {
        private int length = 1;
        private boolean dot = true;

        NumeralText(int len,boolean d){
            super();

            length  = len-1;
            dot = d;

            setFont(new java.awt.Font("dialog", 0, 16));
        }


        //
        //
        //
        public int getValue() {
            int ret = 0;

            String s = getText();
            if(null == s ) return ret;

            try{
                ret = Integer.parseInt(s);
            }
            catch(NumberFormatException e){
                ret = 0;
            }
            return ret;
        }

        //
        //
        //
        protected Document createDefaultModel() {
            return new NumericDocument();
        }

        //
        //
        //
        class NumericDocument extends PlainDocument {

            //
            //
            public void insertString( int offset, String str, AttributeSet a )
                        throws BadLocationException {

                if(length < getLength()) return;

                String validValues = null;

                if(dot){
                    validValues = "0123456789.";
                }
                else {
                    validValues = "0123456789";
                }

                char[] val = str.toCharArray();
                for (int i = 0; i < val.length; i++) {
                    if(validValues.indexOf(val[i]) == -1) return;
                }
                super.insertString( offset, str, a );
            }
        }
    }

    /***************************************************************************
     *
     *       ファイル名を設定するTextField
     *
     ***************************************************************************/
    public class FileText extends JTextField {

        FileText(){
            super();
            setFont(new java.awt.Font("dialog", 0, 16));
        }

        //
        //
        protected Document createDefaultModel() {
            return new NumericDocument();
        }

        //
        //
        class NumericDocument extends PlainDocument {
            String validValues = "0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz.-_";

            //
            //
            public void insertString( int offset, String str, AttributeSet a )
                    throws BadLocationException {

                char[] val = str.toCharArray();
                for (int i = 0; i < val.length; i++) {
                    if(validValues.indexOf(val[i]) == -1) return;
                }
                super.insertString( offset, str, a );
            }
        }
    }

}
