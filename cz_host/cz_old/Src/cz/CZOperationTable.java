package cz;

import java.awt.Component;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.awt.event.WindowAdapter;
import java.awt.event.WindowEvent;
import java.text.DecimalFormat;
import java.util.Locale;
import java.util.Vector;

import javax.swing.JButton;
import javax.swing.JDialog;
import javax.swing.JLabel;
import javax.swing.JOptionPane;
import javax.swing.JScrollPane;
import javax.swing.JTable;
import javax.swing.JTextField;
import javax.swing.JTree;
import javax.swing.ListSelectionModel;
import javax.swing.Timer;
import javax.swing.event.TreeSelectionEvent;
import javax.swing.event.TreeSelectionListener;
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
import javax.swing.tree.DefaultMutableTreeNode;
import javax.swing.tree.DefaultTreeSelectionModel;
import javax.swing.tree.TreePath;
import javax.swing.tree.TreeSelectionModel;

import czclass.CZResult;

/*
 *   操業定数変更Window 
 * @author  (KPK Co.,Ltd.)
 * @version 1.0 (2003/04/01)
 * Update 2013.10.21 他基地参照機能 (@20131021)
 */
public class CZOperationTable extends JDialog {

    private boolean haita_flg           = false;    //排他
    private float send_data[];                      //

    private JButton     send_button     = null;     //実行
    private JButton     cancel_button   = null;     //項目修正
    private JButton     item_button     = null;     //終了

    private TText       op_name         = null;     //設定者

    private DefaultMutableTreeNode  top = null;     //操業定数Tree_Node
    private JTree           tree        = null;     //操業定数Tree
    private JScrollPane     treepanel   = null;     //操業定数Tree_Panel


    private OPTable         tbl         = null;     //操業定数Table
    private JScrollPane     tblpanel    = null;     //操業定数Table_Panel

    private ItemWin         item_win    = null;     //項目設定用Window

    private int current_lag = -1;
    private int current_mid = -1;

/******************************/
    private CloseAlermWin closeAlermWin_    = null;
	public Timer       t                   = null;
	public Timer       at                  = null;
	public Timer       att                 = null;
	public Timer       tcnt                = null;
	
	private int         tcount              = CZSystemDefine.ALERM_DIALOG_CLOSE_TIME/1000;
/******************************/

    //ここからコンストラクタ
    //
    //
    CZOperationTable(){
        super();

        setTitle("操業定数設定");
        setSize(890,480);
        setResizable(false);
        setModal(true);

        addWindowListener(new WindowAdapter(){
            public void windowClosing(WindowEvent e){
                winClose(e);
            }
        });

        getContentPane().setLayout(null);
        // 他基地参照機能    @20131021
        if(CZSystemDefine.REFERENCE_RUN == CZSystem.getRunLevel()){
            getContentPane().setBackground(CZSystemDefine.DEFAULT_REFERENCE_BACKGROUND_COL);
        }else{
            getContentPane().setBackground(CZSystemDefine.DEFAULT_BACKGROUND_COL);
        }

        JLabel  lab = new JLabel("設定者",JLabel.CENTER);
        lab.setBounds(20, 400, 100, 24);
        lab.setLocale(new Locale("ja","JP"));
        lab.setFont(new java.awt.Font("dialog", 0, 16));
        lab.setBorder(new Flush3DBorder());
        lab.setForeground(java.awt.Color.black);
        getContentPane().add(lab);

        op_name = new TText();
        op_name.setBounds(120, 400, 140, 24);
        getContentPane().add(op_name);

        send_button = new JButton("実  行");
        send_button.setBounds(260, 400, 100, 24);
        send_button.setLocale(new Locale("ja","JP"));
        send_button.setFont(new java.awt.Font("dialog", 0, 18));
        send_button.setBorder(new Flush3DBorder());
        send_button.setForeground(java.awt.Color.black);
        send_button.addActionListener(new SendButton());
        getContentPane().add(send_button);

        item_button = new JButton("項目修正");
        item_button.setBounds(380, 400, 100, 24);
        item_button.setLocale(new Locale("ja","JP"));
        item_button.setFont(new java.awt.Font("dialog", 0, 18));
        item_button.setBorder(new Flush3DBorder());
        item_button.setForeground(java.awt.Color.black);
        item_button.addActionListener(new ItemButton());
        getContentPane().add(item_button);

        cancel_button = new JButton("終  了");
        cancel_button.setBounds(760, 400, 100, 24);
        cancel_button.setLocale(new Locale("ja","JP"));
        cancel_button.setFont(new java.awt.Font("dialog", 0, 18));
        cancel_button.setBorder(new Flush3DBorder());
        cancel_button.setForeground(java.awt.Color.black);
        cancel_button.addActionListener(new CancelButton());
        getContentPane().add(cancel_button);

        top = new DefaultMutableTreeNode("操業定数");

        for(int i = 0 ; ; i++){ // for 1
            CZSystemOpTbLag largename = CZSystem.getOpTbLag(i);
            if(null == largename) break;

            DefaultMutableTreeNode large = new DefaultMutableTreeNode(largename.k_name.trim());
            top.add(large);

            for(int j = 0 ; ; j++){ // for 2
                CZSystemOpTbMid middlename = CZSystem.getOpTbMid(i,j);
                if(null == middlename) break;
                Node middle = new Node(middlename.k_name.trim(),middlename);
                large.add(middle);
            } // for 2 end
        } // for 1 end

        DefaultTreeSelectionModel model = new DefaultTreeSelectionModel();
        model.setSelectionMode(TreeSelectionModel.SINGLE_TREE_SELECTION);

        tree = new JTree(top);
        tree.setSelectionModel(model);
        tree.addTreeSelectionListener(new TreeSelect());

        treepanel = new JScrollPane(tree);
        treepanel.setBounds(20, 20, 240, 354);
        treepanel.setBorder(new Flush3DBorder());
        treepanel.setForeground(java.awt.Color.black);
        getContentPane().add(treepanel);

        tbl = new OPTable();
        JTableHeader tabHead = tbl.getTableHeader();
        tabHead.setReorderingAllowed(false);

        tblpanel = new JScrollPane(tbl);
        tblpanel.setBounds(260, 20, 600, 354);
        getContentPane().add(tblpanel);

        // @20131021 
        if(CZSystemDefine.REFERENCE_RUN == CZSystem.getRunLevel()){
            CZSystem.log("CZOperationTable RUNLEVEL", "RUNLEVEL : " +CZSystem.getRunLevel());
            send_button.setEnabled(false);
        }
        // @20131021

        item_win = new ItemWin();
        item_win.setVisible(false);

/*************************/
        //画面クローズ警告画面
        closeAlermWin_ = new CloseAlermWin();
        closeAlermWin_.setVisible(false);

        if( 0 != CZSystemDefine.TIMER_FLG ){
	        t = new Timer( CZSystemDefine.CT_TABLE_CLOSE_TIME, new AlermWin() );
	        tcnt = new Timer( 1000, new CountDown() );
	        at = new Timer( CZSystemDefine.ALERM_DIALOG_CLOSE_TIME, new CancelClose() );
	        att = new Timer( 10, new HaitaKaihou() );
		}
/*************************/
    }


/**************************/

    /**
     *
     *       画面クローズアラーム用Window
     *
     */
    public class CloseAlermWin extends JDialog {
		
		public JLabel       cnt_lab         = null;
		private JLabel      lab             = null;
		private JButton     cancel_button   = null;
		
	    //
	    // コンストラクタ
	    //
	    CloseAlermWin(){
	        super();

	        setTitle("画面クローズ警告");
	        setSize(400,150);
	        setLocation(600,500);
	        setResizable(false);
	        setModal(true);
	        
	        addWindowListener(new WindowAdapter(){
	            public void windowClosing(WindowEvent e){
	                AlermWinClose(e);
	            }
	        });

	        getContentPane().setLayout(null);
            // 他基地参照機能    @20131021
            if(CZSystemDefine.REFERENCE_RUN == CZSystem.getRunLevel()){
                getContentPane().setBackground(CZSystemDefine.DEFAULT_REFERENCE_BACKGROUND_COL);
            }else{
                getContentPane().setBackground(CZSystemDefine.DEFAULT_BACKGROUND_COL);
            }

	        cancel_button = new JButton("了解");
	        cancel_button.setBounds(150, 60, 100, 24);
	        cancel_button.setLocale(new Locale("ja","JP"));
	        cancel_button.setFont(new java.awt.Font("dialog", 0, 18));
	        cancel_button.setBorder(new Flush3DBorder());
	        cancel_button.setForeground(java.awt.Color.black);
	        cancel_button.addActionListener(new AlermClose());
	        getContentPane().add(cancel_button);
	        
			cnt_lab = new JLabel("");
			cnt_lab.setBounds(70, 10, 30, 30);
			cnt_lab.setLocale(new Locale("ja","JP"));
			cnt_lab.setFont(new java.awt.Font("dialog", 0, 18));
//			cnt_lab.setBorder(new Flush3DBorder());
			cnt_lab.setForeground(java.awt.Color.black);
			getContentPane().add(cnt_lab);	        

			lab = new JLabel("秒後に画面を閉じます");
			lab.setBounds(100, 10, 250, 30);
			lab.setLocale(new Locale("ja","JP"));
			lab.setFont(new java.awt.Font("dialog", 0, 18));
			lab.setForeground(java.awt.Color.black);
			getContentPane().add(lab);	        
	    }
	}

	/********************************
	*
	* カウントダウン
	*
	********************************/
	class CountDown implements ActionListener{
		public void actionPerformed( ActionEvent a ){
			
			tcount = tcount - 1;
			
			Integer i = new Integer( tcount );
			
			CZSystem.log("CZControlTable","アラーム画面 クローズまで" + i + "秒");
			
			closeAlermWin_.cnt_lab.setText( i.toString() );
		}
	}

	/********************************
	*
	* アラーム画面
	*
	********************************/
	class AlermWin implements ActionListener{
		public void actionPerformed( ActionEvent a ){
			t.stop();
			at.restart();
			tcount = CZSystemDefine.ALERM_DIALOG_CLOSE_TIME/1000;
			tcnt.restart();
			CZSystem.log("CZOperationTable","アラーム画面");
			closeAlermWin_.cnt_lab.setText("");
			closeAlermWin_.setVisible(true);
			
		}
	}


	/********************************
	*
	* アラーム画面クローズ
	*
	********************************/
	class AlermClose implements ActionListener{
		public void actionPerformed( ActionEvent a ){
			at.stop();
			tcnt.stop();
			t.restart();
			CZSystem.log("CZOperationTable","アラーム画面オープンリスタート（アラーム了解）");
			CZSystem.log("CZOperationTable","アラーム画面クローズ");
			closeAlermWin_.setVisible(false);
		}
	}

	/********************************
	*
	* 画面クローズ
	*
	********************************/
	class CancelClose implements ActionListener{
		public void actionPerformed( ActionEvent a ){
			at.stop();
			tcnt.stop();
			t.stop();
			item_win.setVisible(false);
			closeAlermWin_.setVisible(false);
			setVisible(false);
			att.restart();
		}
	}

	/********************************
	*
	* 排他開放
	*
	********************************/
	class HaitaKaihou implements ActionListener{
		public void actionPerformed( ActionEvent a ){
			putHaita();
			att.stop();
		}
	}

    //
    // アラーム画面クローズ
    //
    private void AlermWinClose(WindowEvent e){
        CZSystem.log("CZOperationTable","AlermWinClose() " + e);
			at.stop();
			tcnt.stop();
			t.restart();
			CZSystem.log("CZOperationTable","アラーム画面オープンリスタート（×）");
			CZSystem.log("CZOperationTable","アラーム画面クローズ");
    }

	public boolean timerStart(){
		at.stop();
		t.restart();
		CZSystem.log("CZOperationTable","アラーム画面オープンリスタート（メニュー）");
		CZSystem.log("CZOperationTable","デフォルト設定");
	
		return true;
	}
	
/**************************/

    //
    //
    //
    private boolean selectData(int lag , int mid ){

//@@        CZSystem.log("CZOperationTable","selectData (" + lag + " : " + mid + ")");

        Vector data = null;
        data = CZSystem.getOpTb(lag,mid);

        if(null == data) return false;
        
        tbl = null;
        tbl = new OPTable(data);
        JTableHeader tabHead = tbl.getTableHeader();
        tabHead.setReorderingAllowed(false);

        tblpanel.setViewportView(tbl);

        current_lag = lag;
        current_mid = mid;

        return true;
    }

    //
    //
    //
    private boolean setSendStatus(){
//@@        CZSystem.log("CZOperationTable","setSendStatus (" + current_lag + " : " + current_mid + ")");

        if(1 > op_name.getText().length()){
//@@            CZSystem.log("CZOperationTable","setSendStatus() Table Op Name Error !!");
            Object msg[] = {"操業定数更新",
                                "設定者を入力してくださ！！",
                                ""};
            errorMsg(msg);
            return false;
        }

        if(1 > current_lag){
//@@            CZSystem.log("CZOperationTable","setSendStatus() Table Data Lag Error !!");
            Object msg[] = {"操業定数更新",
                                "選択を見直してください！！",
                                ""};
            errorMsg(msg);
            return false;
        }

        if(1 > current_mid){
//@@            CZSystem.log("CZOperationTable","setSendStatus() Table Data Mid Error !!");
            Object msg[] = {"操業定数更新",
                                "選択を見直してください！！",
                                ""};
            errorMsg(msg);
            return false;
        }

        if(tbl.isEditing()){
//@@            CZSystem.log("CZOperationTable","setSendStatus() Table Data Edit Error !!");
            Object msg[] = {"操業定数更新",
                                "入力を見直してください！！",
                                "設定中項目有り"};
            errorMsg(msg);
            return false;
        }

        if(!tblCheck()){
//@@            CZSystem.log("CZOperationTable","setSendStatus() Table Data Error !!");
            Object msg[] = {"操業定数更新",
                                "入力を見直してください！！",
                                "上下限"};
            errorMsg(msg);
            return false;
        }

//@@        CZSystem.log("CZOperationTable","setSendStatus() Table Data OK !!!!");


        int count = tbl.getRowCount();
        send_data = new float[count];
        for(int row = 0 ; row < count ; row++){
            Float   val  = (Float)tbl.getValueAt(row,6);
            send_data[row] = val.floatValue();
        }
        return true;
    }

    //
    //
    //
    private boolean tblCheck(){

        int count = tbl.getRowCount();
        CZSystem.log("CZOperationTable","tblCheck(" + count + ")");

        for(int row = 0 ; row < count ; row++){

            if(Float.class != tbl.getValueAt(row, 2).getClass()){
                return false;
            }
            if(Float.class != tbl.getValueAt(row, 3).getClass()){
                return false;
            }
            if(Float.class != tbl.getValueAt(row, 6).getClass()){
                return false;
            }

            Float   min  = (Float)tbl.getValueAt(row,2);
            Float   max  = (Float)tbl.getValueAt(row,3);
            Float   val  = (Float)tbl.getValueAt(row,6);

            if((min.floatValue() <= val.floatValue()) &&
                    (max.floatValue() >= val.floatValue())){
                continue ;
            }
            else {
                CZSystem.log("CZOperationTable","tblCheck *** error ***( min=" + min + ": max="+ max +" val=" + val + ")");
                return false;
            }
        } // for end    

        return true;
    }



    //
    // 排他取得要求
    //
    private boolean getHaita(){

        // 既に取ってる場合
        if(haita_flg) return true;

        String ro = CZSystem.getRoName();

        CZEventCL event = null;

        CZSystemQueue   que = new CZSystemQueue(20);
        CZEventAdapter  adp = new CZEventAdapter(que);
        CZEventSender.addCZEventListener(adp);

        boolean ret = CZSystem.CZGetWorkingExclusion(ro);

        haita_flg = false;

        if(!ret){
            CZEventSender.removeCZEventListener(adp);
            return false;
        }

        while(true){
            try{
//@@                CZSystem.log("CZOperationTable","getHaita(1)");
                event = (CZEventCL)que.waitObject();

//@@                CZSystem.log("CZOperationTable","getHaita(2)");
                // 排他応答以外
                if(event.getEvent() != CZEventCL.OT_GET_HAITA) continue;
//@@                CZSystem.log("CZOperationTable","getHaita(3)");

                CZResult ev = (CZResult)event.getObject();

//@@                CZSystem.log("CZOperationTable","getHaita(4)");
                // 違う炉の場合
                if(!ro.equals(ev.getRoban())) continue;

//@@                CZSystem.log("CZOperationTable","getHaita(5)");

                // 排他取得失敗
                if(0 != ev.getStatus()){
//@@                    CZSystem.log("CZOperationTable","getHaita(6)");
                    CZEventSender.removeCZEventListener(adp);

                    CZSystemSysMsg msg = new CZSystemSysMsg();
                    msg.no = -1;
                    msg.message = CZSystem.getDateTime() + " 操業定数排他取得失敗 [" + ev.getStatus() + "]";
                    CZSystem.sysMessage(msg);
                    return false;
                }

//@@                CZSystem.log("CZOperationTable","getHaita(7)");
                CZEventSender.removeCZEventListener(adp);
                haita_flg = true;

                CZSystemSysMsg msg = new CZSystemSysMsg();
                msg.no = 0;
                msg.message = CZSystem.getDateTime() + " 操業定数排他取得成功 [" + ev.getStatus() + "]";
                CZSystem.sysMessage(msg);
                return true;
            }
            catch(Exception e){
                CZSystem.handleException(e);
            }
        } //while end
    }


    //
    // 排他開放要求
    //
    private boolean putHaita(){

        String ro = CZSystem.getRoName();
        // 常に開放する様に変更     01.03.27    
        boolean ret = CZSystem.CZPutWorkingExclusion(ro);
        haita_flg = false;
//@@        CZSystem.log("CZOperationTable","putHaita(排他開放要求 2)");

        return true;
    }


    //
    //
    //
    private void winClose(WindowEvent e){
        if( 0 != CZSystemDefine.TIMER_FLG ){
			t.stop();
	        at.stop();
	        att.stop();
	        tcnt.stop();
		}
        CZSystem.log("CZOperationTable","winClose() " + e);
        putHaita();
    }


    //
    //
    //
    public boolean setDefault(){

//@@        CZSystem.log("CZOperationTable","setDefault() Start");

        // @20131021 他基地参照機能
        if(CZSystemDefine.REFERENCE_RUN != CZSystem.getRunLevel()){  // 参照のみの場合、排他処理は実行しない

        if(!getHaita()){
            Object msg[] = {"操業定数排他取得",
                    "制御盤、他の端末で",
                    "修正中です"};
            errorMsg(msg);
        }

        send_button.setEnabled(haita_flg);

        }  // @20131021

        current_lag = -1;
        current_mid = -1;

        tbl = null;
        tbl = new OPTable();
        JTableHeader tabHead = tbl.getTableHeader();
        tabHead.setReorderingAllowed(false);
        tblpanel.setViewportView(tbl);

        op_name.setText("");
//@@        CZSystem.log("CZOperationTable","setDefault() Exit");

        return true;
    }


        //
        // メッセージの表示
        //
    private boolean errorMsg(Object msg[]){
        JOptionPane.showMessageDialog(null,msg,
            "操業定数入力エラー",
            JOptionPane.ERROR_MESSAGE);
        return true;
    }


    /*
    * SendButton ActionListener
    *
    *
    */
    class SendButton implements ActionListener {
        public void actionPerformed(ActionEvent ev){

            if(setSendStatus()){
//@@                CZSystem.log("CZOperationTable SendButton","----->"+ send_data);

                //Send
                CZSystem.CZWorkingTableExchnage(op_name.getText(),
                            current_lag,current_mid,
                            send_data);
            }
            return ;
        }
    }


    /*
    * ItemButton ActionListener
    *
    *
    */
    class ItemButton implements ActionListener {
        public void actionPerformed(ActionEvent ev){

        if( 0 != CZSystemDefine.TIMER_FLG ){
			at.stop();
			t.restart();
			CZSystem.log("CZOperationTable","アラーム画面オープンリスタート（項目修正）");
			CZSystem.log("CZOperationTable","デフォルト設定");
		}
		
            int  item_no = tbl.getSelectedRow();

            if(0 > item_no) return;
            item_no++;

//@@            CZSystem.log("CZOperationTable ItemButton","[" + current_lag + "][" + current_mid + "][" + item_no + "]");

            if(null == item_win) return;

            item_win.setDefault(current_lag,current_mid,item_no);   

            item_win.setVisible(true);

            return ;
        }
    }

    /*
    *
    *
    *
    */
    class CancelButton implements ActionListener {
        public void actionPerformed(ActionEvent ev){
			if( 0 != CZSystemDefine.TIMER_FLG ){
				at.stop();
				tcnt.stop();
				t.stop();
			}
            setVisible(false);
            putHaita();
        }
    }


    /*
    *
    *
    *
    */
    class Node extends DefaultMutableTreeNode {
        private CZSystemOpTbMid data = null;
        Node(String name,CZSystemOpTbMid dat){
            super(name);
            data = dat;
        }

        public CZSystemOpTbMid getData(){
            return data;
        }
    }

    /*
    *
    *
    *
    */
    class TreeSelect implements TreeSelectionListener {
        public void valueChanged(TreeSelectionEvent ev){

            TreePath path  = ev.getPath();
            int      count = path.getPathCount();
            DefaultMutableTreeNode node  =  
                (DefaultMutableTreeNode)path.getLastPathComponent();

            if(node.isLeaf()){

                Node n = (Node)node;
                CZSystemOpTbMid dat = n.getData();
//@@                CZSystem.log("CZOperationTable","TreeSelect  [" +
//@@                    dat.k_no1 + "][" + dat.k_no2 + "][" + dat.k_name + "]");

                if(!selectData(dat.k_no1,dat.k_no2)) setDefault();
            }
            else {
                setDefault();
            }
        }
    }


    /*
    *
    *   操業定数テーブルクラス
    *
    */


    public class OPTable extends JTable {

        private OPTblMdl model = null;

        OPTable(){
            super();

            setName("OPTable");
            setBounds(0, 0, 200, 200);
            setAutoCreateColumnsFromModel(true);
            setSelectionMode(ListSelectionModel.SINGLE_SELECTION);
            setLocale(new Locale("ja","JP"));
            setFont(new java.awt.Font("dialog", 0, 12));
            setRowHeight(17);
        }

        OPTable(Vector data){
            super();

            try{
                setName("OPTable");
                setBounds(0, 0, 200, 200);
                setAutoCreateColumnsFromModel(true);
                setSelectionMode(ListSelectionModel.SINGLE_SELECTION);

                setLocale(new Locale("ja","JP"));
                setFont(new java.awt.Font("dialog", 0, 12));
                setRowHeight(17);

                model = new OPTblMdl(data);

                setModel(model);


                DefaultTableColumnModel cmdl = (DefaultTableColumnModel)getColumnModel();
                TableColumn  colum = null;
                OPTblRenderer ren  = null;

                //#
                colum = cmdl.getColumn(0);
                colum.setMaxWidth(40);
                colum.setMinWidth(40);
                colum.setWidth(40);
                ren = new OPTblRenderer();
                ren.setHorizontalAlignment(ren.RIGHT);
                colum.setCellRenderer(ren);

                //項目名    
                colum = cmdl.getColumn(1);
                colum.setMaxWidth(220);
                colum.setMinWidth(220);
                colum.setWidth(220);
                ren = new OPTblRenderer();
                colum.setCellRenderer(ren);

                //Min
                colum = cmdl.getColumn(2);
                colum.setMaxWidth(60);
                colum.setMinWidth(60);
                colum.setWidth(60);
                ren = new OPTblRenderer();
                ren.setHorizontalAlignment(ren.RIGHT);
                colum.setCellRenderer(ren);

                //Max
                colum = cmdl.getColumn(3);
                colum.setMaxWidth(60);
                colum.setMinWidth(60);
                colum.setWidth(60);
                ren = new OPTblRenderer();
                ren.setHorizontalAlignment(ren.RIGHT);
                colum.setCellRenderer(ren);

                //桁
                colum = cmdl.getColumn(4);
                colum.setMaxWidth(30);
                colum.setMinWidth(30);
                colum.setWidth(30);
                ren = new OPTblRenderer();
                ren.setHorizontalAlignment(ren.RIGHT);
                colum.setCellRenderer(ren);

                //単位
                colum = cmdl.getColumn(5);
                colum.setMaxWidth(100);
                colum.setMinWidth(100);
                colum.setWidth(100);
                ren = new OPTblRenderer();
                colum.setCellRenderer(ren);


                //値
                colum = cmdl.getColumn(6);
                colum.setMaxWidth(70);
                colum.setMinWidth(70);
                colum.setWidth(70);
                ren = new OPTblRenderer();
                ren.setHorizontalAlignment(ren.RIGHT);
                colum.setCellRenderer(ren);
            }
            catch (Throwable e) {
                CZSystem.handleException(e);
            }
        }
    }

    /*
    *
    *   操業定数テーブルクラス：モデル
    *
    */

    public class  OPTblMdl extends AbstractTableModel {

        final   int TBL_COL = 7;
        private int TBL_ROW = 0;

        private Object data[][];

        final String[] names = {"#", "項    目",
                    "Min","Max","桁","単位","値"};

        OPTblMdl(Vector dat){
            super();

            if(null == dat) return;

            TBL_ROW = dat.size();
            data = new Object[TBL_ROW][TBL_COL];

            try{

                for(int i = 0 ; i < TBL_ROW ; i++){ 

                    CZSystemOpTb d = (CZSystemOpTb)dat.elementAt(i);
//@@                    System.out.println("***** OPTblMdl:kNo1=" + d.k_no1 + " :kNo2=" + d.k_no2 + " :kNo=" + d.k_no);
                    CZSystemOpTbSml k = CZSystem.getOpTbSml(d.k_no1,d.k_no2,d.k_no);

                    data[i][0] = new Integer(d.k_no);
                    data[i][6] = new Float(d.k_val);

                    if(null != k){
                        if(null == k.k_name) data[i][1] = new String("NULL");
                        else if(1 > k.k_name.length()) data[i][1] = new String("NULL");
                        else data[i][1] = k.k_name.trim();

                        data[i][2] = new Float(k.n_min);
                        data[i][3] = new Float(k.n_max);
                        data[i][4] = new Integer(k.keta);

                        if(null == k.t_name) data[i][5] = new String("NULL");
                        else if(1 > k.t_name.length()) data[i][5] = new String("NULL");
                        else data[i][5] = k.t_name.trim();
                    }
                    else {
                        data[i][1] = new String("################################");
                        data[i][2] = new String("#.#####");
                        data[i][3] = new String("#####.#####");
                        data[i][5] = new String("######");
                        data[i][6] = new String("##");
                    }
                }
            }
            catch (Throwable e) {
                CZSystem.handleException(e);
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
            if(6 == col) return true;
            return false;
        }

        public void setValueAt(Object aValue, int row, int column){ 
            data[row][column] = aValue;
        }
    }


    /*
    *
    *   操業定数テーブルクラス：レンダラー
    *
    *
    */

    class OPTblRenderer extends DefaultTableCellRenderer {

        OPTblRenderer(){
        super();
            setLocale(new Locale("ja","JP"));
            setFont(new java.awt.Font("dialog", 0, 12));
        }

        public Component getTableCellRendererComponent( JTable table,
                                                        Object value,
                                                        boolean isSelected,
                                                        boolean hasFocus,   
                                                        int row,int column){
            

            if(6 != column){                                
                super.getTableCellRendererComponent(table,
                                                    value,
                                                    isSelected,
                                                    hasFocus,   
                                                    row,column);
                return(this);           
            }

            if(Float.class != table.getValueAt(row, 2).getClass()){
                super.getTableCellRendererComponent(table,
                                                    value,
                                                    isSelected,
                                                    hasFocus,   
                                                    row,column);
                return(this);           
            }

            if(Float.class != table.getValueAt(row, 3).getClass()){
                super.getTableCellRendererComponent(table,
                                                    value,
                                                    isSelected,
                                                    hasFocus,   
                                                    row,column);
                return(this);           
            }

            if(Integer.class != table.getValueAt(row, 4).getClass()){
                super.getTableCellRendererComponent(table,
                                                    value,
                                                    isSelected,
                                                    hasFocus,   
                                                    row,column);
                return(this);           
            }


            Float   min  = (Float)table.getValueAt(row,2);
            Float   max  = (Float)table.getValueAt(row,3);
            Integer keta = (Integer)table.getValueAt(row,4);
            Float   val  = (Float)table.getValueAt(row,6);

            if(null == val) {
                val = new Float(0.0f);
            }

            DecimalFormat format = null;
            StringBuffer  buff = new StringBuffer();

            if(1 > keta.intValue()){
                format = new DecimalFormat("0");
            }
            else {
                buff.append("0.");
                for(int i = 0 ; i < keta.intValue() ; i++){
                    buff.append("0");
                }
                format = new DecimalFormat(buff.toString());
            }

            Float new_val = new Float(format.format(val));

            super.getTableCellRendererComponent(table,
                                                format.format(new_val.floatValue()),
                                                isSelected,
                                                hasFocus,   
                                                row,column);

            table.setValueAt(new_val,row,column);

            if((min.floatValue() <= new_val.floatValue()) &&
                (max.floatValue() >= new_val.floatValue())){
                setForeground(java.awt.Color.blue);
            }
            else {
                setForeground(java.awt.Color.red);
            }
            return(this);           
        }
    }


    /*
    *
    *       設定者を入力するTextField
    *
    */
    public class TText extends JTextField {

        TText(){
            super();
            setFont(new java.awt.Font("dialog", 0, 16));
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
            String validValues = "0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ-/.:";

            //
            //
            public void insertString( int offset, String str, AttributeSet a )
            throws BadLocationException {


                if(29 < getLength()) return;
                char[] val = str.toCharArray();
                for (int i = 0; i < val.length; i++) {
                    if(validValues.indexOf(val[i]) == -1) return;
                }

                super.insertString( offset, str, a );
            }

            //
            //
            public void remove(int offs, int len)
                throws BadLocationException {
                super.remove(offs,len);
            }
        }
    }


    /*
    *
    *   項目設定用Window
    *
    */

    public class ItemWin extends JDialog {

        private JButton item_send_button    = null;
        private JButton item_cancel_button  = null;
        private TText   item_op_name        = null;

        private ItemText   item_name    = null;
        private MinMaxText item_min = null;
        private MinMaxText item_max = null;
        private DigitText  item_digit   = null;
        private UnitText   item_unit    = null;

        private String  sendOp      = null;
        private String  sendName    = null;
        private float   sendMin;
        private float   sendMax;
        private int sendDigit;
        private String  sendUnit    = null;


        private int item_lag = -1;
        private int item_mid = -1;
        private int item_no  = -1;

        //ここからコンストラクタ
        //
        //
        ItemWin(){
            super();

            setTitle("操業定数項目設定");
            setSize(765,150);
            setResizable(false);
            setModal(true);

            addWindowListener(new WindowAdapter(){
                public void windowClosing(WindowEvent e){
                    AlermWinClose(e);
                }
            });

            getContentPane().setLayout(null);
            // 他基地参照機能    @20131021
            if(CZSystemDefine.REFERENCE_RUN == CZSystem.getRunLevel()){
                getContentPane().setBackground(CZSystemDefine.DEFAULT_REFERENCE_BACKGROUND_COL);
            }else{
                getContentPane().setBackground(CZSystemDefine.DEFAULT_BACKGROUND_COL);
            }


            JLabel lab;
            
            lab = new JLabel("項                目",JLabel.CENTER);
            lab.setBounds(20, 20, 300, 24);
            lab.setLocale(new Locale("ja","JP"));
            lab.setFont(new java.awt.Font("dialog", 0, 16));
            lab.setBorder(new Flush3DBorder());
            lab.setForeground(java.awt.Color.black);
            getContentPane().add(lab);

            lab = new JLabel("Ｍｉｎ",JLabel.CENTER);
            lab.setBounds(330, 20, 100, 24);
            lab.setLocale(new Locale("ja","JP"));
            lab.setFont(new java.awt.Font("dialog", 0, 16));
            lab.setBorder(new Flush3DBorder());
            lab.setForeground(java.awt.Color.black);
            getContentPane().add(lab);
            
            lab = new JLabel("Ｍａｘ",JLabel.CENTER);
            lab.setBounds(440, 20, 100, 24);
            lab.setLocale(new Locale("ja","JP"));
            lab.setFont(new java.awt.Font("dialog", 0, 16));
            lab.setBorder(new Flush3DBorder());
            lab.setForeground(java.awt.Color.black);
            getContentPane().add(lab);

            lab = new JLabel("桁",JLabel.CENTER);
            lab.setBounds(550, 20, 25, 24);
            lab.setLocale(new Locale("ja","JP"));
            lab.setFont(new java.awt.Font("dialog", 0, 16));
            lab.setBorder(new Flush3DBorder());
            lab.setForeground(java.awt.Color.black);
            getContentPane().add(lab);

            lab = new JLabel("単        位",JLabel.CENTER);
            lab.setBounds(585, 20, 150, 24);
            lab.setLocale(new Locale("ja","JP"));
            lab.setFont(new java.awt.Font("dialog", 0, 16));
            lab.setBorder(new Flush3DBorder());
            lab.setForeground(java.awt.Color.black);
            getContentPane().add(lab);

            //
            item_name = new ItemText();
            item_name.setBounds(20, 44, 300, 24);
            item_name.setLocale(new Locale("ja","JP"));
            item_name.setFont(new java.awt.Font("dialog", 0, 16));
            item_name.setBorder(new Flush3DBorder());
            item_name.setForeground(java.awt.Color.black);
            getContentPane().add(item_name);

            item_min = new MinMaxText();
            item_min.setBounds(330, 44, 100, 24);
            item_min.setLocale(new Locale("ja","JP"));
            item_min.setFont(new java.awt.Font("dialog", 0, 16));
            item_min.setBorder(new Flush3DBorder());
            item_min.setForeground(java.awt.Color.black);
            getContentPane().add(item_min);

            item_max = new MinMaxText();
            item_max.setBounds(440, 44, 100, 24);
            item_max.setLocale(new Locale("ja","JP"));
            item_max.setFont(new java.awt.Font("dialog", 0, 16));
            item_max.setBorder(new Flush3DBorder());
            item_max.setForeground(java.awt.Color.black);
            getContentPane().add(item_max);

            item_digit = new DigitText();
            item_digit.setBounds(550, 44, 25, 24);
            item_digit.setLocale(new Locale("ja","JP"));
            item_digit.setFont(new java.awt.Font("dialog", 0, 16));
            item_digit.setBorder(new Flush3DBorder());
            item_digit.setForeground(java.awt.Color.black);
            getContentPane().add(item_digit);


            item_unit = new UnitText();
            item_unit.setBounds(585, 44, 150, 24);
            item_unit.setLocale(new Locale("ja","JP"));
            item_unit.setFont(new java.awt.Font("dialog", 0, 16));
            item_unit.setBorder(new Flush3DBorder());
            item_unit.setForeground(java.awt.Color.black);
            getContentPane().add(item_unit);

            //
            lab = new JLabel("設定者",JLabel.CENTER);
            lab.setBounds(20, 80, 100, 24);
            lab.setLocale(new Locale("ja","JP"));
            lab.setFont(new java.awt.Font("dialog", 0, 16));
            lab.setBorder(new Flush3DBorder());
            lab.setForeground(java.awt.Color.black);
            getContentPane().add(lab);

            item_op_name = new TText();
            item_op_name.setBounds(120, 80, 140, 24);
            getContentPane().add(item_op_name);

            item_send_button = new JButton("実  行");
            item_send_button.setBounds(260, 80, 100, 24);
            item_send_button.setLocale(new Locale("ja","JP"));
            item_send_button.setFont(new java.awt.Font("dialog", 0, 18));
            item_send_button.setBorder(new Flush3DBorder());
            item_send_button.setForeground(java.awt.Color.black);
            item_send_button.addActionListener(new ItemSendButton());
            getContentPane().add(item_send_button);

            item_cancel_button = new JButton("終  了");
            item_cancel_button.setBounds(635, 80, 100, 24);
            item_cancel_button.setLocale(new Locale("ja","JP"));
            item_cancel_button.setFont(new java.awt.Font("dialog", 0, 18));
            item_cancel_button.setBorder(new Flush3DBorder());
            item_cancel_button.setForeground(java.awt.Color.black);
            item_cancel_button.addActionListener(new ItemCancelButton());
            getContentPane().add(item_cancel_button);

        }


        //
        //
        //
        public boolean setDefault(int lag,int mid,int item){
            if(1 > lag) return false;
            if(1 > mid) return false;
            if(1 > item) return false;

            item_lag = lag;
            item_mid = mid;
            item_no  = item;

            int row = item_no -1;

            String  s1  = (String)tbl.getValueAt(row,1);
            item_name.setText(s1);

            Float   s2  = (Float)tbl.getValueAt(row,2);
            item_min.setText(s2.toString());

            Float   s3  = (Float)tbl.getValueAt(row,3);
            item_max.setText(s3.toString());

            Integer s4  = (Integer)tbl.getValueAt(row,4);
            item_digit.setText(s4.toString());

            String  s5  = (String)tbl.getValueAt(row,5);
            item_unit.setText(s5);

            item_op_name.setText("");

            if(CZSystemDefine.ADMIN_RUN == CZSystem.getRunLevel()){
                item_send_button.setEnabled(true);
            }
            else{
                item_send_button.setEnabled(false);
            }
            return true;
        }


        //
        //
        //
        private boolean setItemSendStatus(){

            sendOp = item_op_name.getText();
            if(1 > sendOp.length()){
                return false;
            }

            sendName = item_name.getText();
            if(1 > sendName.length()){
                return false;
            }

            sendUnit = item_unit.getText();
            if(1 > sendUnit.length()){
                return false;
            }

            try{
                sendMin   = Float.parseFloat(item_min.getText());
                sendMax   = Float.parseFloat(item_max.getText());
                sendDigit = Integer.parseInt(item_digit.getText());
            }
            catch (Exception e){
                return false;
            }

            if(sendMin >= sendMax) return false;
            if(0 > sendDigit) return false;
            return true;
        }


        /*
        *
        *
        *
        */
        class ItemSendButton implements ActionListener {
            public void actionPerformed(ActionEvent ev){

                if(!setItemSendStatus()){
                    Object msg[] = {"操業定数項目更新",
                                    "設定者、項目、Min、Max、桁を",
                                    "見直してください"};
                    errorMsg(msg);
                    return;
                }

//@@                CZSystem.log("CZOperationTable","ItemSendButton ----->[" +
//@@                                        item_lag  + "][" + item_mid + "][" + item_no + "][" +
//@@                                        sendOp    + "][" + sendName + "][" +    
//@@                                        sendMin   + "][" + sendMax  + "][" +    
//@@                                        sendDigit + "][" + sendUnit + "]");

                //Send
                if(!CZSystem.CZWorkingNameExchnage(sendOp, item_lag, item_mid, item_no,
                               sendName, sendUnit, sendMin, sendMax, sendDigit)){
                    Object msg[] = {"操業定数項目更新",
                                    "更新が失敗しました",
                                    "管理者に問い合わせてください"};
                    errorMsg(msg);
                    return;
                }
                return ;
            }
        }


        /*
        *
        *
        *
        */
        class ItemCancelButton implements ActionListener {
            public void actionPerformed(ActionEvent ev){
				if( 0 != CZSystemDefine.TIMER_FLG ){
					at.stop();
					t.restart();
					CZSystem.log("CZOperationTable","アラーム画面オープンリスタート（修正項目終了）");
				}
                setVisible(false);
            }
        }


        /*
        *       操業定数：項目名を入力するTextField
        */
        public class ItemText extends JTextField {

            ItemText(){
                 super();
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

                    String tmp = new String(getText(0,getLength()) + str);
                    byte   b[];

                    try{
                        b = tmp.getBytes("SJIS");
                    }
                    catch(Exception e){
//@@                        CZSystem.log("CZOperationTable","ItemText [" + e + "]");
                        return;
                    }

//@@                    CZSystem.log("CZOperationTable","ItemText [" + tmp + "][" + b + "][" + b.length + "]");

                    if(32 < b.length) return;
//                    if(50 < b.length) return;
                    super.insertString( offset, str, a );
                }

                //
                //
                public void remove(int offs, int len)
                    throws BadLocationException {
                    super.remove(offs,len);
                }
            }
        }


        /*
        *       操業定数：ＭｉｎＭａｘを入力するTextField
        */
        public class MinMaxText extends JTextField {

            MinMaxText(){
                super();
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
                String validValues = "0123456789.-";

                //
                //
                public void insertString( int offset, String str, AttributeSet a )
                throws BadLocationException {
            
                    if(9 < getLength()) return;
                    char[] val = str.toCharArray();
                    for (int i = 0; i < val.length; i++) {
                        if(validValues.indexOf(val[i]) == -1) return;
                    }

                    super.insertString( offset, str, a );
                }

                //
                //
                public void remove(int offs, int len)
                    throws BadLocationException {
                    super.remove(offs,len);
                }
            }
        }

        /*
        *       操業定数：桁を入力するTextField
        */
        public class DigitText extends JTextField {

            DigitText(){
                super();
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
                String validValues = "0123456";

                //
                //
                public void insertString( int offset, String str, AttributeSet a )
                throws BadLocationException {
            
                    if(0 < getLength()) return;
                    char[] val = str.toCharArray();
                    for (int i = 0; i < val.length; i++) {
                        if(validValues.indexOf(val[i]) == -1) return;
                    }
                    super.insertString( offset, str, a );
                }

                //
                //
                public void remove(int offs, int len)
                    throws BadLocationException {
                    super.remove(offs,len);
                }

            }
        }


        /*
        *       操業定数：単位を入力するTextField
        */
        public class UnitText extends JTextField {

            UnitText(){
                super();
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

                    String tmp = new String(getText(0,getLength()) + str);
                    byte   b[];

                    try{
                        b = tmp.getBytes("SJIS");
                    }
                    catch(Exception e){
//@@                        CZSystem.log("CZOperationTable","ItemText [" + e + "]");
                        return;
                    }

//@@                    CZSystem.log("CZOperationTable","ItemText [" + tmp + "][" + b + "][" + b.length + "]");

                    if(16 < b.length) return;
                    super.insertString( offset, str, a );
                }

                //
                //
                public void remove(int offs, int len)
                    throws BadLocationException {
                    super.remove(offs,len);
                }
           }
        }
    }
}
