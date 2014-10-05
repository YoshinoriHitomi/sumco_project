package cz;

import java.awt.Component;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.util.Locale;
import java.util.Vector;

import javax.swing.JButton;
import javax.swing.JDialog;
import javax.swing.JLabel;
import javax.swing.JOptionPane;
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

/***********************************************************
 *
 *   エラー表示Window 
 * @author  (KPK Co.,Ltd.)
 * @version 1.0 (2003/04/01)
 *
 ***********************************************************/
public class CZErrorMsgWin extends JDialog {

    private final   int DEFAULT_DAY = 1;
    private final   int  M_ERROR_MAST_SG = 128;		/* 2006.4.13 y.k */
    private final   int  M_ERROR_MAST_AP = 1000;		/* 2006.4.13 y.k */

    private Vector      err_list    = null;
    private Vector      err_list_mast_ro = null;

    private JScrollPane err_panel   = null;
    private ErrorTable  err_tbl     = null;

    private DayText     day_text    = null;

    // ---------- コンストラクタ ---------------------------
    //
    CZErrorMsgWin(){
        super();

        setTitle("システムエラー");
        setSize(910,455);
        setResizable(false);
        setModal(true);

        getContentPane().setLayout(null);
        // 他基地参照機能    @20131021
        if(CZSystemDefine.REFERENCE_RUN == CZSystem.getRunLevel()){
            getContentPane().setBackground(CZSystemDefine.DEFAULT_REFERENCE_BACKGROUND_COL);
        }else{
            getContentPane().setBackground(CZSystemDefine.DEFAULT_BACKGROUND_COL);
        }

        err_panel = new JScrollPane();
        err_panel.setBounds(20, 20, 870, 360);
        getContentPane().add(err_panel);

        JLabel label = new JLabel("表示日数",JLabel.CENTER);
        label.setBounds(20, 390, 100, 24);
        label.setLocale(new Locale("ja","JP"));
        label.setFont(new java.awt.Font("dialog", 0, 16));
        label.setBorder(new Flush3DBorder());
        label.setForeground(java.awt.Color.black);
        getContentPane().add(label);

        day_text = new DayText();
        day_text.setBounds(120, 390, 40, 24);
        getContentPane().add(day_text);

        JButton day_button = new JButton("再読込み");
        day_button.setBounds(160, 390, 100, 24);
        day_button.setLocale(new Locale("ja","JP"));
        day_button.setFont(new java.awt.Font("dialog", 0, 18));
        day_button.setBorder(new Flush3DBorder());
        day_button.setForeground(java.awt.Color.black);
        day_button.addActionListener(new ModifyButton());
        getContentPane().add(day_button);

    }


    //
    //
    //
    public boolean setDefault(){
//@@        CZSystem.log("CZErrorMsgWin","setDefault()");
        day_text.setText(String.valueOf(DEFAULT_DAY));
        String db = CZSystem.getDBName();

		/* 2006.04.13 y.k */
		err_list_mast_ro = CZSystem.errorMessageRead2(db);

        err_list = CZSystem.getRoError(db,DEFAULT_DAY);
        ErrorTable t = new ErrorTable();
        JTableHeader tabHead = t.getTableHeader();
        tabHead.setReorderingAllowed(false);
        err_panel.setViewportView(t);

        return true;
    }


    /*******************************************************
     *
     *       エラー実績一覧
     *
     *******************************************************/
    class ErrorTable extends JTable {

        private ErrorTblMdl model = null;

        // ---------- コンストラクタ -----------------------
        ErrorTable(){
            super();

            try{
                setName("ErrorTable");
                setAutoCreateColumnsFromModel(true);
                setSelectionMode(ListSelectionModel.SINGLE_SELECTION);
                setLocale(new Locale("ja","JP"));
                setFont(new java.awt.Font("dialog", 0, 12));
                setRowHeight(17);
                if(null == err_list) return;
                if(1 > err_list.size()) return;
                model = new ErrorTblMdl(err_list.size());
                setModel(model);
                DefaultTableColumnModel cmdl = (DefaultTableColumnModel)getColumnModel();
                TableColumn  colum = null;
                ErrorTblRenderer ren   = null;
                // No
                colum = cmdl.getColumn(0);
                colum.setMaxWidth(40);
                colum.setMinWidth(40);
                colum.setWidth(40);
                // 発生日時
                colum = cmdl.getColumn(1);
                colum.setMaxWidth(162);
                colum.setMinWidth(162);
                colum.setWidth(162);
                ren = new ErrorTblRenderer();
                ren.setHorizontalAlignment(ren.CENTER);
                colum.setCellRenderer(ren);
                // バッチNo
                colum = cmdl.getColumn(2);
                colum.setMaxWidth(130);
                colum.setMinWidth(130);
                colum.setWidth(130);
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
                // エラーNo
                colum = cmdl.getColumn(5);
                colum.setMaxWidth(60);
                colum.setMinWidth(60);
                colum.setWidth(60);
                // メッセージ
                colum = cmdl.getColumn(6);
                colum.setMaxWidth(250);
                colum.setMinWidth(250);
                colum.setWidth(250);
                // 状態
                colum = cmdl.getColumn(7);
                colum.setMaxWidth(60);
                colum.setMinWidth(60);
                colum.setWidth(60);

                ren = new ErrorTblRenderer();
                ren.setHorizontalAlignment(ren.CENTER);
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
        }

        //
        //
        //
        public void setData(int gr,int tbl){

//@@            CZSystem.log("CZErrorMsgWin","setData [" + gr + "][" + tbl + "]");
        }


        /***************************************************
         *
         *      エラー実績一覧：モデル
         *
         ***************************************************/
        public class ErrorTblMdl extends AbstractTableModel {

            private int TBL_ROW     = 0;
            final   int TBL_COL     = 8;

            final String[] names = {" # "          , "発生日時" ,   
                    "Bt"           , "プロセス" ,
                    "プロセス時間" , "E-Code",
                    "メッセージ"   , "状態" };

            private Object  data[][];

            // ---------- コンストラクタ -------------------
            ErrorTblMdl(int max){
                super();
                TBL_ROW = max;
                if(1 > TBL_ROW) return;
                data = new Object[TBL_ROW][TBL_COL];
                for(int i = 0 ; i < TBL_ROW ; i++){
                    CZSystemErr err = (CZSystemErr)err_list.elementAt(i);

                    // No
                    data[i][0] = new Integer(i+1);
                    // 発生日時
                    data[i][1] = err.o_time;
                    // バッチNo
                    if ( null != err.batch ) {
                        data[i][2] = err.batch;
                    } else {
                        data[i][2] = new String("");
                    }
                    // プロセス
                    data[i][3] = CZSystem.getProcName(err.p_no);
                    // プロセス時間
                    data[i][4] = new Integer(err.p_time);
                    // エラーNo
                    data[i][5] = new Integer(err.e_no);
                    // メッセージ
					/* 2006.04.13 y.k */
					if ((err.e_no <= M_ERROR_MAST_SG) || (err.e_no >= M_ERROR_MAST_AP)) {
	                    CZSystemErrMsg e_msg = CZSystem.getErrMsg(err.e_no);
	                    data[i][6] = e_msg.message;
    	            } else {
					    CZSystemErrMsg e_msg = CZSystem.getErrMsg2(err.e_no, err_list_mast_ro);
    	                data[i][6] = e_msg.message;
					}
                    // 状態
                    data[i][7] = getState(err.flg_error);
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

            public String getState(int no){
                String ret = null;

                switch(no){
                    case 0 : return "復旧";

                    case 1 : return "発生";
                }
                return "不明";
            }
        } // ErrorTblMdl


        /***************************************************
         *
         *       エラー実績一覧：レンダラー
         *
         ***************************************************/
        class ErrorTblRenderer extends DefaultTableCellRenderer {

            ErrorTblRenderer(){
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

                if(7 != column) return(this);
                String val = (String)value;
                if(val.equals("復旧")) setForeground(java.awt.Color.blue);
                if(val.equals("発生")) setForeground(java.awt.Color.red);
                if(val.equals("不明")) setForeground(java.awt.Color.black);
                return(this);
            }
        } // ErrorTblRenderer
    } // ErrorTable


    /*******************************************************
     *
     *       表示日数を入力するTextField
     *
     *******************************************************/
    public class DayText extends JTextField {

        DayText(){
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
            String validValues = "0123456789";

            //
            //
            public void insertString( int offset, String str, AttributeSet a )
                                throws BadLocationException {

                if(2 < getLength()) return;
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
    } // DayText

    /*******************************************************
     *
     *
     *
     *******************************************************/
    class ModifyButton implements ActionListener {
        public void actionPerformed(ActionEvent ev){
            String  val = null;
            int day = 0;
            val = day_text.getText();
            if(null == val){
                day_text.setText(String.valueOf(DEFAULT_DAY));
                return;
            }
            if(1 > val.length()){
                day_text.setText(String.valueOf(DEFAULT_DAY));
                return;
            }
            try{
                day = Integer.parseInt(val);
                if(1 > day){
                    day_text.setText(String.valueOf(DEFAULT_DAY));
                    return;
                }
                String db = CZSystem.getDBName();

                int cnt = CZSystem.getRoErrorCount(db,day);
                int errMax = CZSystem.getErrorMax();

                if (errMax < cnt) {
                    Object msg[] = {"炉エラー表示",
                              "データ件数が " + errMax + "件 を超えています表示日数を見直してください！！"};

                    JOptionPane.showMessageDialog(null,
                                                  msg,"エラー",
                                                  JOptionPane.ERROR_MESSAGE);

                } else {
                    err_list = CZSystem.getRoError(db,day);
                    ErrorTable t = new ErrorTable();
                    JTableHeader tabHead = t.getTableHeader();
                    tabHead.setReorderingAllowed(false);
                    err_panel.setViewportView(t);
                }
            }
            catch(Exception e){
                day_text.setText(String.valueOf(DEFAULT_DAY));
                return;
            }
            return;
        }
    } // ModifyButton
}
