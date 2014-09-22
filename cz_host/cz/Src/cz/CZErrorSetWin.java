package cz;

import java.awt.Component;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.awt.event.WindowAdapter;
import java.awt.event.WindowEvent;
import java.util.Locale;
import java.util.Vector;

import javax.swing.DefaultCellEditor;
import javax.swing.JButton;
import javax.swing.JComboBox;
import javax.swing.JDialog;
import javax.swing.JLabel;
import javax.swing.JScrollPane;
import javax.swing.JTable;
import javax.swing.JTextField;
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

import czclass.CZParamErrorDefine;
import czclass.CZParamErrorMsgDefine;

/***********************************************************
 *
 *   エラー項目変更Window 
 * @author  (KPK Co.,Ltd.)
 * @version 1.0 (2003/04/01)
 *
 ***********************************************************/
public class CZErrorSetWin extends JDialog {

    private Vector     err_list         = null;
    private ErrorTable err_tbl          = null;

    private JButton    send_button      = null;
    private JButton    cancel_button    = null;
    private TText      op_name          = null;

    private Vector     send_data        = null;
    private Vector     errmsg_data      = null;

    // ---------- コンストラクタ ---------------------------
    //
    CZErrorSetWin(){
        super();

        setTitle("エラー項目定義");
//        setSize(1120,455);
        setSize(1130+60,455);
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
        lab.setBounds(20, 390, 100, 24);
        lab.setLocale(new Locale("ja","JP"));
        lab.setFont(new java.awt.Font("dialog", 0, 16));
        lab.setBorder(new Flush3DBorder());
        lab.setForeground(java.awt.Color.black);
        getContentPane().add(lab);

        op_name = new TText();
        op_name.setBounds(120, 390, 140, 24);
        getContentPane().add(op_name);

        send_button = new JButton("実  行");
        send_button.setBounds(260, 390, 100, 24);
        send_button.setLocale(new Locale("ja","JP"));
        send_button.setFont(new java.awt.Font("dialog", 0, 18));
        send_button.setBorder(new Flush3DBorder());
        send_button.setForeground(java.awt.Color.black);
        send_button.addActionListener(new SendButton());
        getContentPane().add(send_button);

        cancel_button = new JButton("終  了");
        cancel_button.setBounds(990, 390, 100, 24);
        cancel_button.setLocale(new Locale("ja","JP"));
        cancel_button.setFont(new java.awt.Font("dialog", 0, 18));
        cancel_button.setBorder(new Flush3DBorder());
        cancel_button.setForeground(java.awt.Color.black);
        cancel_button.addActionListener(new CancelButton());
        getContentPane().add(cancel_button);

        err_tbl = new ErrorTable();
        JTableHeader tabHead = err_tbl.getTableHeader();
        tabHead.setReorderingAllowed(false);

        JScrollPane panel = new JScrollPane(err_tbl);
//        panel.setBounds(20, 20, 1070, 356);
        panel.setBounds(20, 20, 1080+60, 356);
        getContentPane().add(panel);

    }


    //
    //
    //
    private void winClose(WindowEvent e){
        CZSystem.log("CZErrorSetWin","winClose() " + e);
        err_tbl.clearSelection();
    } 


    //
    //
    //
    public boolean setDefault(){
//@@        CZSystem.log("CZErrorSetWin","setDefault() ");

        if(CZSystemDefine.ADMIN_RUN == CZSystem.getRunLevel()){
            send_button.setEnabled(true);
        }
        else{
            send_button.setEnabled(false);
        }

        err_tbl.clearSelection();
        op_name.setText("");

        err_list = CZSystem.getErrTitle();
        if(null == err_list){
            CZSystem.exit(-1,"CZErrorSetWin setDefault getErrTitle() Error");
        }

        for(int i = 0 ; i < err_list.size() ; i++){
            CZSystemErrName err = (CZSystemErrName)err_list.elementAt(i);
//@@            System.out.println("CZSystemErrName [" + err.e_no + "][" + err.e_name + "][" + err.process + "]");

            // #
            int j = 0;
            Integer no = new Integer(err.e_no);
            err_tbl.setValueAt(no, i, j);

            // 項目名
            j++;
            err_tbl.setValueAt(err.e_name, i, j);
            // プロセス
            for(int k = 0 ; k <= CZSystemDefine.END ; k++){
                j++;
                int ret = getProcFlag(k , err.process);
                switch(ret){
                    case 1: err_tbl.setValueAt(new String("○"), i, j);
                    break;

                    default : err_tbl.setValueAt(new String(" "), i, j);
                    break;
                }
            }

            // 立上り
            j++;
            switch(err.edge){
                case 0: err_tbl.setValueAt(new String("立上"), i, j);
                break;

                case 1: err_tbl.setValueAt(new String("立下"), i, j);
                break;

                case -1: err_tbl.setValueAt(new String("   "), i, j);
                break;

                default : err_tbl.setValueAt(new String("   "), i, j);
                break;
            }

            // レディー移行
            j++;
            switch(err.ready){
                case 1: err_tbl.setValueAt(new String("移行"), i, j);
                break;

                default : err_tbl.setValueAt(new String("   "), i, j);
                break;
            }
            // 区分 
            j++;
            switch(err.kubun){
                case 1: err_tbl.setValueAt(new String("警報"), i, j);
                break;

                case 2: err_tbl.setValueAt(new String("通知"), i, j);
                break;

                case 3: err_tbl.setValueAt(new String("警報+通知"), i, j);
                break;

                default : err_tbl.setValueAt(new String("   "), i, j);
                break;
            }
            // 表示場所
            j++;
            switch(err.basho){
                case 1: err_tbl.setValueAt(new String("PC"), i, j);
                break;

                case 2: err_tbl.setValueAt(new String("PLC"), i, j);
                break;

                case 3: err_tbl.setValueAt(new String("PC+PLC"), i, j);
                break;

                default : err_tbl.setValueAt(new String("   "), i, j);
                break;
            }
            // 画面表示
            j++;
            switch(err.umu){
                case 1: err_tbl.setValueAt(new String("有"), i, j);
                break;

                default : err_tbl.setValueAt(new String(" "), i, j);
                break;
            }
            // ブザー１回
            j++;
            switch(err.buzzer1){
                case 1: err_tbl.setValueAt(new String("有"), i, j);
                break;

                default : err_tbl.setValueAt(new String(" "), i, j);
                break;
            }
            // ブザー連続
            j++;
            switch(err.buzzer){
                case 1: err_tbl.setValueAt(new String("有"), i, j);
                break;

                default : err_tbl.setValueAt(new String(" "), i, j);
                break;
            }
            // エラー処理
            j++;
            switch(err.error_umu){
                case 1: err_tbl.setValueAt(new String("有"), i, j);
                break;

                default : err_tbl.setValueAt(new String(" "), i, j);
                break;
            }
            // 自己復帰
            j++;
            switch(err.fukkyu){
                case 1: err_tbl.setValueAt(new String("有"), i, j);
                break;

                default : err_tbl.setValueAt(new String(" "), i, j);
                break;
            }
            // 監視表示 2006.09.28 tuika 
            j++;
            switch(err.hyoji){
                case 1: err_tbl.setValueAt(new String("表示"), i, j);
                break;

                default : err_tbl.setValueAt(new String(" "), i, j);
                break;
            }
        } // for end
        err_tbl.repaint();
        return true;
    }



    //
    //
    //
	@SuppressWarnings("unchecked")
    private boolean setSendStatus(){

        if(1 > op_name.getText().length()){
            return false;
        }
        send_data = new Vector(CZSystemDefine.ERROR_MAX);
        errmsg_data = new Vector(130);
        for(int i = 0 ; i < err_list.size() ; i++){
            CZParamErrorDefine d = new CZParamErrorDefine();
            CZParamErrorMsgDefine md = new CZParamErrorMsgDefine();
//@@            CZSystem.log("CZErrorSetWin","setSendStatus [" + i + "]");
            // 項目名No
            int pos = 0;
            Integer _no = (Integer)err_tbl.getValueAt(i,pos);
            if(null == _no ) return false;
            int no = _no.intValue();
CZSystem.log("CZErrorSetWin","no : " + no);
            if(1 > no) return false;
            if(CZSystemDefine.ERROR_MAX < no) return false;
            d.setErrorNo(no);
            md.setErrorNo(no);
//@@            CZSystem.log("CZErrorSetWin","setSendStatus [" + i + "][" + pos + "]");
            // 項目名
            pos++;
            String name = (String)err_tbl.getValueAt(i,pos);
            if(null == name ) return false;
CZSystem.log("CZErrorSetWin","name : " + name);
            if(1 > name.length()) return false;
            d.setErrorName(name);
            md.setErrorName(name);
//@@            CZSystem.log("CZErrorSetWin","setSendStatus [" + i + "][" + pos + "]");
            // プロセス
            int proc[] = new int[CZSystemDefine.END+1];
            for(int j = 0 ; j <= CZSystemDefine.END ; j++){
                pos++;
                String maru =  (String)err_tbl.getValueAt(i,pos);
                if(null == maru) return false;
                if(maru.equals("○")) proc[j] = 1;
                else                  proc[j] = 0;
            }
            d.setProcess(proc);
//@@            CZSystem.log("CZErrorSetWin","setSendStatus [" + i + "][" + pos + "]");
            // 立上り
            pos++;
            String tachi = (String)err_tbl.getValueAt(i,pos);
            if(null == tachi ) return false;
            if(tachi.equals("立上"))      d.setEdge(0);
            else if(tachi.equals("立下")) d.setEdge(1);
            else                          d.setEdge(-1);
//@@            CZSystem.log("CZErrorSetWin","setSendStatus [" + i + "][" + pos + "]");
            // レディー移行
            pos++;
            String rdy = (String)err_tbl.getValueAt(i,pos);
            if(null == rdy ) return false;
            if(rdy.equals("移行"))      d.setReady(1);
            else                        d.setReady(0);
//@@            CZSystem.log("CZErrorSetWin","setSendStatus [" + i + "][" + pos + "]");
            // 区分
            pos++;
            String kubun = (String)err_tbl.getValueAt(i,pos);
            if(null == kubun ) return false;
            if(kubun.equals("警報"))           d.setKubun(1);
            else if(kubun.equals("通知"))      d.setKubun(2);
            else if(kubun.equals("警報+通知")) d.setKubun(3);
            else                               d.setKubun(0);
//@@            CZSystem.log("CZErrorSetWin","setSendStatus [" + i + "][" + pos + "]");
            // 表示場所
            pos++;
            String basho = (String)err_tbl.getValueAt(i,pos);
            if(null == basho ) return false;
            if(basho.equals("PC"))          d.setBasho(1);
            else if(basho.equals("PLC"))    d.setBasho(2);
            else if(basho.equals("PC+PLC")) d.setBasho(3);
            else                            d.setBasho(0);
//@@            CZSystem.log("CZErrorSetWin","setSendStatus [" + i + "][" + pos + "]");
            // 画面表示
            pos++;
            String gamen = (String)err_tbl.getValueAt(i,pos);
            if(null == gamen ) return false;
            if(gamen.equals("有")) d.setDisp_umu(1);
            else                   d.setDisp_umu(0);
//@@            CZSystem.log("CZErrorSetWin","setSendStatus [" + i + "][" + pos + "]");
            // ブザー１回
            pos++;
            String bu1 = (String)err_tbl.getValueAt(i,pos);
            if(null == bu1 ) return false;
            if(bu1.equals("有"))   d.setBuzzer1(1);
            else                   d.setBuzzer1(0);
//@@            CZSystem.log("CZErrorSetWin","setSendStatus [" + i + "][" + pos + "]");
            // ブザー連続
            pos++;
            String bu = (String)err_tbl.getValueAt(i,pos);
            if(null == bu ) return false;
            if(bu.equals("有"))    d.setBuzzer(1);
            else                   d.setBuzzer(0);
//@@            CZSystem.log("CZErrorSetWin","setSendStatus [" + i + "][" + pos + "]");
            // エラー処理
            pos++;
            String err = (String)err_tbl.getValueAt(i,pos);
            if(null == err ) return false;
            if(err.equals("有"))   d.setError_umu(1);
            else                   d.setError_umu(0);
//@@            CZSystem.log("CZErrorSetWin","setSendStatus [" + i + "][" + pos + "]");
            // 自己復旧
            pos++;
            String jiko = (String)err_tbl.getValueAt(i,pos);
            if(null == jiko ) return false;
            if(jiko.equals("有"))  d.setFukkyu(1);
            else                   d.setFukkyu(0);
//@@            CZSystem.log("CZErrorSetWin","setSendStatus [" + i + "][" + pos + "]");

            // 監視表示 2006.09.28 tuika
            pos++;
            String M_Disp = (String)err_tbl.getValueAt(i,pos);
            if(null == M_Disp ) return false;
            if(M_Disp.equals("表示"))
			{
			  d.setHyoji(1);
//CZSystem.log("CZErrorSetWin","set 表示 [" + i + "][" + M_Disp + "]");
            }
			else
			{
                   d.setHyoji(0);
//CZSystem.log("CZErrorSetWin","set 未表示 [" + i + "]");
			}

            send_data.addElement(d);    
            errmsg_data.addElement(md);    
        } // for end
        return true;
    }


    //
    //
    //
    public int getProcFlag(int proc , int val){

        int shift = proc;
        int mask  = 1;
        int tmp = val >>> shift;
        int ret = mask & tmp;
//@@        System.out.println("getProcFlag Proc[" + proc + "] tmp[" + 
//@@                                                 tmp  + "] ret[" + ret + "]");
        return ret;
    }


    /*******************************************************
     *
     *   エラーテーブル
     *
     *******************************************************/
    class ErrorTable extends JTable {

        private CtTblMdl model = null;

        ErrorTable(){
            super();
            try{
                setName("ErrorTable");
//                setBounds(0, 0, 200, 200);
                setBounds(0, 0, 210+60, 200);
                setAutoCreateColumnsFromModel(true);
                setSelectionMode(ListSelectionModel.SINGLE_SELECTION);
                setLocale(new Locale("ja","JP"));
                setFont(new java.awt.Font("dialog", 0, 12));
                setRowHeight(17);
                model = new CtTblMdl();
                setModel(model);
                DefaultTableColumnModel cmdl = (DefaultTableColumnModel)getColumnModel();
                TableColumn  colum = null;
                // #
                int i = 0;
                colum = cmdl.getColumn(i);
//                colum.setMaxWidth(30);
//                colum.setMinWidth(30);
//                colum.setWidth(30);
                colum.setMaxWidth(40);
                colum.setMinWidth(40);
                colum.setWidth(40);
                // 項目名
                i++;
                colum = cmdl.getColumn(i);
                colum.setMaxWidth(230);
                colum.setMinWidth(230);
                colum.setWidth(230);
                colum.setCellEditor(new ItemCell(new ItemText()));
                // プロセス
                for(int j = 0 ; j <= CZSystemDefine.END ; j++){
                    i++;
                    colum = cmdl.getColumn(i);
                    colum.setMaxWidth(40);
                    colum.setMinWidth(40);
                    colum.setWidth(40);
                    colum.setCellRenderer(new ErrRenderer());
                    colum.setCellEditor(new ErrCell(new ProcComboBox()));
                }
                // 立上り
                i++;
                colum = cmdl.getColumn(i);
                colum.setMaxWidth(60);
                colum.setMinWidth(60);
                colum.setWidth(60);
                colum.setCellRenderer(new ErrRenderer());
                colum.setCellEditor(new ErrCell(new UpComboBox()));
                // READY移行
                i++;
                colum = cmdl.getColumn(i);
                colum.setMaxWidth(60);
                colum.setMinWidth(60);
                colum.setWidth(60);
                colum.setCellRenderer(new ErrRenderer());
                colum.setCellEditor(new ErrCell(new ReadyComboBox()));
                // 区分
                i++;
                colum = cmdl.getColumn(i);
                colum.setMaxWidth(60);
                colum.setMinWidth(60);
                colum.setWidth(60);
                colum.setCellRenderer(new ErrRenderer());
                colum.setCellEditor(new ErrCell(new KubunComboBox()));
                // 表示場所
                i++;
                colum = cmdl.getColumn(i);
                colum.setMaxWidth(60);
                colum.setMinWidth(60);
                colum.setWidth(60);
                colum.setCellRenderer(new ErrRenderer());
                colum.setCellEditor(new ErrCell(new BashoComboBox()));
                // 画面表示
                i++;
                colum = cmdl.getColumn(i);
                colum.setMaxWidth(30);
                colum.setMinWidth(30);
                colum.setWidth(30);
                colum.setCellRenderer(new ErrRenderer());
                colum.setCellEditor(new ErrCell(new UseComboBox()));
                // ブザー一回
                i++;
                colum = cmdl.getColumn(i);
                colum.setMaxWidth(30);
                colum.setMinWidth(30);
                colum.setWidth(30);
                colum.setCellRenderer(new ErrRenderer());
                colum.setCellEditor(new ErrCell(new UseComboBox()));
                // ブザー連続
                i++;
                colum = cmdl.getColumn(i);
                colum.setMaxWidth(30);
                colum.setMinWidth(30);
                colum.setWidth(30);
                colum.setCellRenderer(new ErrRenderer());
                colum.setCellEditor(new ErrCell(new UseComboBox()));
                // エラー処理
                i++;
                colum = cmdl.getColumn(i);
                colum.setMaxWidth(30);
                colum.setMinWidth(30);
                colum.setWidth(30);
                colum.setCellRenderer(new ErrRenderer());
                colum.setCellEditor(new ErrCell(new UseComboBox()));
                // 自己復旧
                i++;
                colum = cmdl.getColumn(i);
                colum.setMaxWidth(30);
                colum.setMinWidth(30);
                colum.setWidth(30);
                colum.setCellRenderer(new ErrRenderer());
                colum.setCellEditor(new ErrCell(new UseComboBox()));
                // 監視表示
                i++;
                colum = cmdl.getColumn(i);
                colum.setMaxWidth(60);
                colum.setMinWidth(60);
                colum.setWidth(60);
                colum.setCellRenderer(new ErrRenderer());
                colum.setCellEditor(new ErrCell(new DispComboBox()));
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
//@@            CZSystem.log("CZErrorSetWin","valueChanged [" + getSelectedRow() + "][" + getSelectedColumn() + "]");
            if(0 > getSelectedRow()) return;
        }
    }

    /*******************************************************
     *
     *       制御テーブルクラス：モデル
     *
     *******************************************************/
    public class CtTblMdl extends AbstractTableModel {

        final   int TBL_COL     = 22;
        private int TBL_ROW     = CZSystemDefine.ERROR_MAX;

        final String[] names = {"#", 
                    "項目名",
                    CZSystem.getProcName(CZSystemDefine.READY) ,
                    CZSystem.getProcName(CZSystemDefine.VAC) ,
                    CZSystem.getProcName(CZSystemDefine.MELT) ,
                    CZSystem.getProcName(CZSystemDefine.DIP) ,
                    CZSystem.getProcName(CZSystemDefine.NECK1) ,
                    CZSystem.getProcName(CZSystemDefine.NECK2) ,
                    CZSystem.getProcName(CZSystemDefine.SHOULDER) ,
                    CZSystem.getProcName(CZSystemDefine.BODY) ,
                    CZSystem.getProcName(CZSystemDefine.TAIL) ,
                    CZSystem.getProcName(CZSystemDefine.END) ,
                    "立上り",
                    "R移行",
                    "区分",
                    "場所",
                    "画面",
                    "B-1",
                    "B-C",
                    "処理",
                    "復旧",
                    "監視表示"};

        private Object  data[][];

        CtTblMdl(){
            super();

            data = new Object[TBL_ROW][TBL_COL];

            for(int i = 0 ; i < TBL_ROW ; i++){
                data[i][0] = new Integer(i+1);
                data[i][1] = new String("123456789012345678901234567890");

                data[i][2] = new String("  ");
                data[i][3] = new String("  ");
                data[i][4] = new String("  ");
                data[i][5] = new String("  ");
                data[i][6] = new String("  ");
                data[i][7] = new String("  ");
                data[i][8] = new String("  ");
                data[i][9] = new String("  ");
                data[i][10] = new String("  ");
                data[i][11] = new String("  ");

                data[i][12] = new String("    ");
                data[i][13] = new String("    ");
                data[i][14] = new String("    ");
                data[i][15] = new String("    ");
                data[i][16] = new String("    ");
                data[i][17] = new String("    ");
                data[i][18] = new String("    ");
                data[i][19] = new String("    ");
                data[i][20] = new String("    ");
                data[i][21] = new String("    ");
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

            if(0 == col) return false;
            return true;
        }

        public void setValueAt(Object aValue, int row, int column){
            data[row][column] = aValue;
        }
    }

    /*******************************************************
     *
     *******************************************************/
    public class ErrCell extends DefaultCellEditor {

        ErrCell(JComboBox box){
            super(box);
        }
    }


    /*******************************************************
     *
     *******************************************************/
    public class ItemCell extends DefaultCellEditor {


        //
        //
        //
        ItemCell(ItemText tx){
            super(tx);
            setClickCountToStart(1);
        }

        //
        //
        //
        public Component getTableCellEditorComponent( JTable table,
                                                    Object value,
                                                    boolean isSelected,
                                                    boolean hasFocus,
                                                    int row,int column){
            Component ret = super.getTableCellEditorComponent(table,value,isSelected,row,column);
            return ret;
        }
    }



    /*******************************************************
     *
     *******************************************************/
    public class ProcComboBox extends JComboBox {

        ProcComboBox(){
            super();
            addItem("  ");
            addItem("○");
        }
    }

    /*******************************************************
     *
     *******************************************************/
    public class UpComboBox extends JComboBox {

        UpComboBox(){
            super();
            addItem("    ");
            addItem("立上");
            addItem("立下");
        }
    }

    /*******************************************************
     *
     *******************************************************/
    public class ReadyComboBox extends JComboBox {

        ReadyComboBox(){
            super();
            addItem("    ");
            addItem("移行");
        }
    }


    /*******************************************************
     *
     *******************************************************/
    public class KubunComboBox extends JComboBox {

        KubunComboBox(){
            super();
            addItem("警報");
            addItem("通知");
            addItem("警報+通知");
        }
    }


    /*******************************************************
     *
     *******************************************************/
    public class BashoComboBox extends JComboBox {

        BashoComboBox(){
            super();
            addItem("PC");
            addItem("PLC");
            addItem("PC+PLC");
        }
    }


    /*******************************************************
     *
     *******************************************************/
    public class UseComboBox extends JComboBox {

        UseComboBox(){
            super();
            addItem("  ");
            addItem("有");
        }
    }

    /*******************************************************
     *　2006.09.28　tuika
     *******************************************************/
    public class DispComboBox extends JComboBox {

        DispComboBox(){
            super();
            addItem("  ");
            addItem("表示");
        }
    }

    /*******************************************************
     *       項目名を入力するTextField
     *******************************************************/
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
                    CZSystem.log("CZErrorSetWin","ItemText [" + e + "]");
                    return;
                }

//@@                CZSystem.log("CZErrorSetWin","ItemText [" + tmp + "][" + b + "][" + b.length + "]");
                if(30 < b.length) return;
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


    /*******************************************************
     *
     *******************************************************/
    class ItemRenderer extends DefaultTableCellRenderer {

        ItemRenderer(){
            super();
            setLocale(new Locale("ja","JP"));
            setFont(new java.awt.Font("dialog", 0, 12));
            setHorizontalAlignment(LEFT);
        }

        public Component getTableCellRendererComponent( JTable table,
                                                    Object value,
                                                    boolean isSelected,
                                                    boolean hasFocus,
                                                    int row,int column){
            String s1 = (String)value;
            String s2 = (String)err_tbl.getValueAt(row,column);
//@@            CZSystem.log("CZErrorSetWin","Object[" + s1 + "] ValueAt[" + s2 + "]"); 
            super.getTableCellRendererComponent(table,
                                                value,
                                                isSelected,
                                                hasFocus,
                                                row,column);
            return(this);
        }
    }


    /*******************************************************
     *
     *******************************************************/
    class ErrRenderer extends DefaultTableCellRenderer {

        ErrRenderer(){
            super();
            setLocale(new Locale("ja","JP"));
            setFont(new java.awt.Font("dialog", 0, 12));
            setHorizontalAlignment(CENTER);
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


    /*******************************************************
     *       設定者を入力するTextField
     *******************************************************/
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

    /*******************************************************
     *
     *******************************************************/
    class SendButton implements ActionListener {
        public void actionPerformed(ActionEvent ev){

            boolean ret = false;
            boolean ret2 = false;
            if(setSendStatus()){
//@@                CZSystem.log("CZErrorSetWin","SendButton TRUE");
                //Send
                ret = CZSystem.CZErrorDefineSend(op_name.getText(),send_data);
                ret2 = CZSystem.CZErrorMsgDefineSend(op_name.getText(),errmsg_data);
            }
            else {
//@@                CZSystem.log("CZErrorSetWin","SendButton FALSE");
            }
//@@            CZSystem.log("CZErrorSetWin","SendButton() [" + ret + "]"); 
            return ;
        }
    }


    /*******************************************************
     *
     *******************************************************/
    class CancelButton implements ActionListener {
        public void actionPerformed(ActionEvent ev){
            err_tbl.clearSelection();
            setVisible(false);
        }
    }
}
