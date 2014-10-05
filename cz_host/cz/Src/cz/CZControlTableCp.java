package cz;

import java.awt.Rectangle;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.awt.event.WindowAdapter;
import java.awt.event.WindowEvent;
import java.util.Locale;
import java.util.Vector;

import javax.swing.JButton;
import javax.swing.JComboBox;
import javax.swing.JDialog;
import javax.swing.JLabel;
import javax.swing.JOptionPane;
import javax.swing.JScrollPane;
import javax.swing.JTable;
import javax.swing.JTextField;
import javax.swing.JViewport;								//2011.04.12 Y.K add
import javax.swing.ListSelectionModel;
import javax.swing.event.ListSelectionEvent;
import javax.swing.plaf.metal.MetalBorders.Flush3DBorder;
import javax.swing.table.AbstractTableModel;
import javax.swing.table.DefaultTableColumnModel;
import javax.swing.table.JTableHeader;						//2011.04.12 Y.K add
import javax.swing.table.TableColumn;
import javax.swing.text.AttributeSet;
import javax.swing.text.BadLocationException;
import javax.swing.text.Document;
import javax.swing.text.PlainDocument;

/***********************************************************
 *
 *   制御テーブルコピー用Window
 * @author  (KPK Co.,Ltd.)
 * @version 1.0 (2003/04/01)
 *
 ***********************************************************/
public class CZControlTableCp {

//    private RoAllCopy   ro_all_win  = null;
//    private GroupCopy   group_win   = null;
//    private RecipeCopy  recipe_win  = null;
//    private TableCopy   table_win   = null;

//    private T6LagCopy  t6LagWin_    = null;
//    private T6MidCopy  t6MidWin_    = null;
//    private T6ItemCopy t6ItemWin_   = null;

    public RoAllCopy   ro_all_win  = null;
    public GroupCopy   group_win   = null;
    public RecipeCopy  recipe_win  = null;
    public TableCopy   table_win   = null;

    public T6LagCopy  t6LagWin_    = null;
    public T6MidCopy  t6MidWin_    = null;
    public T6ItemCopy t6ItemWin_   = null;

    private boolean    haita_flg    = false;

    //
    // ---------- コンストラクタ ---------------------------
    //
    CZControlTableCp(){
        super();

        ro_all_win = new RoAllCopy();
        group_win  = new GroupCopy();
        recipe_win = new RecipeCopy();
        table_win  = new TableCopy();

        t6LagWin_  = new T6LagCopy();
        t6MidWin_  = new T6MidCopy();
        t6ItemWin_ = new T6ItemCopy();

    }


    //
    // 排他取得要求
    //
    private boolean getHaita(int idx){

        String ro = CZSystem.getRoName(idx);

        CZSystem.log("CZControlTableCp getHaita","ro:" + ro);

        // 同じ炉の場合無条件にtrue
        if(ro.equals(CZSystem.getRoName())){
            haita_flg = true;
            return true;
        }

        //他炉の場合の処理
        boolean ret = CZSystem.CZGetControlExclusion(ro);
        haita_flg = false;
        if(!ret) return false;
//@@        CZSystem.log("CZControlTableCp getHaita","1");
        haita_flg = true;
        return true;
    }

    //
    // 排他開放要求
    //
    private boolean putHaita(int idx){

        String ro = CZSystem.getRoName(idx);

        CZSystem.log("CZControlTableCp putHaita","1");
        if(ro.equals(CZSystem.getRoName())){
            haita_flg = false;
            CZSystem.log("CZControlTableCp putHaita","2");
            return true;
        }

//@@        CZSystem.log("CZControlTableCp putHaita","3");

        // 常に開放する様に変更         01.03.27
        boolean ret = CZSystem.CZPutControlExclusion(ro);
        haita_flg = false;
//@@        CZSystem.log("CZControlTableCp putHaita","5");
        return true;
    }

    //
    // 炉全コピー
    //
    public boolean roAllCopy(){

        CZSystem.log("CZControlTableCp","roAllCopy ");
        boolean ret = ro_all_win.setDefault();
        if(ret) ro_all_win.setVisible(true);
        return true;
    }

    //
    // グループコピー
    //
    public boolean groupCopy(int grp_no,String grp){

        CZSystem.log("CZControlTableCp","groupCopy grp_no[" + grp_no +"] grp[" + grp + "]");
        boolean ret = group_win.setDefault(grp_no,grp);
        if(ret) group_win.setVisible(true);
        return true;
    }

    //
    // レシピーコピー
    //
    public boolean recipeCopy(int grp_no,String grp,int rec_no,String title){

        CZSystem.log("CZControlTableCp","recipeCopy grp_no[" + grp_no +"] grp[" + grp +
               "] rec_no[" + rec_no + "] title[" + title +"]");
        boolean ret = recipe_win.setDefault(grp_no,grp,rec_no,title);
        if(ret) recipe_win.setVisible(true);
        return true;
    }

    //
    // テーブルコピー
    //
    public boolean tableCopy(int grp_no,String grp,int rec_no,String title,
                int[] table_no,String[] table){

        for(int i = 0 ; i < table_no.length ; i++) {
            CZSystem.log("CZControlTableCp tableCopy","actionPerformed [" + i +
                "][" + table_no[i] + "][" + table[i] + "]");
		}
//        CZSystem.log("CZControlTableCp","tableCopy grp_no[" + grp_no +"] grp[" + grp +
//               "] rec_no[" + rec_no + "] title[" + title +
//               "] table_no[" + table_no + "] table[" + table + "]");

////2011.04.12 Y.K start
//        boolean ret = table_win.setDefault(grp_no,grp,rec_no,title,table_no,table);
        boolean ret = table_win.setDefault(grp_no,grp,rec_no,title,table_no,table);
////2011.04.12 Y.K end
        if(ret) table_win.setVisible(true);
        return true;
    }



    //
    // 項目コピー
    //
    public boolean itemCoy(int grp_no,String grp,int rec_no,String title,
                int table_no,String table,int item_no,String item_name){

        CZSystem.log("CZControlTableCp","itemCoy ");
/*@@@@
        boolean ret = itemWin_.setDefault(grp_no, grp, rec_no, title,
                 table_no, table, item_no, item_name);
        if(ret) itemWin_.setVisible(true);
@@@@*/
        return true;
    }

    //
    // Ｔ６大項目コピー
    //
    public boolean t6LagCopy(   int grpNo, String grp,
                                int recNo, String recTitle,
                                int lagNo, String lagName){

        CZSystem.log("CZControlTableCp","t6MidCopy grpNo[" + grpNo +"] grp[" + grp +
               "] recNo[" + recNo + "] recTitle[" + recTitle +
               "] lagNo[" + lagNo + "] lagName[" + lagName + "]");

        boolean ret = t6LagWin_.setDefault(grpNo, grp, recNo, recTitle, lagNo, lagName);
        if(ret) t6LagWin_.setVisible(true);

        return true;
    }

    //
    // Ｔ６中項目コピー
    //
    public boolean t6MidCopy(int grpNo, String grp, int recNo, String recTitle,
                int lagNo, String lagName, int midNo, String midName){

        CZSystem.log("CZControlTableCp","t6MidCopy grpNo[" + grpNo +"] grp[" + grp +
               "] recNo[" + recNo + "] recTitle[" + recTitle +
               "] lagNo[" + lagNo + "] lagName[" + lagName +
               "] midNo[" + midNo + "] midName[" + midName + "]");
        boolean ret = t6MidWin_.setDefault(grpNo, grp, recNo, recTitle, lagNo, lagName, midNo, midName);
        if(ret) t6MidWin_.setVisible(true);
        return true;
    }

    //
    // Ｔ６項目コピー
    //
    public boolean t6ItemCopy(int grpNo, String grp, int recNo, String recTitle,
                int lagNo, String lagName, int midNo, String midName,int itemNo,String itemName){

        CZSystem.log("CZControlTableCp","t6ItemCopy grpNo[" + grpNo +"] grp[" + grp +
               "] recNo[" + recNo + "] recTitle[" + recTitle +
               "] lagNo[" + lagNo + "] lagName[" + lagName +
               "] midNo[" + midNo + "] midName[" + midName +
               "] itemNo[" + itemNo + "] itemName[" + itemName + "]");

        boolean ret = t6ItemWin_.setDefault( grpNo, grp, recNo, recTitle,
                lagNo, lagName, midNo, midName, itemNo, itemName);
        if(ret) t6ItemWin_.setVisible(true);

        return true;
    }



    //
    //  炉全コピー用画面
    //
    public class RoAllCopy extends JDialog {

        private JLabel  src_ro_name = null;

        private RoNo    dst_ro_name = null;

        private JButton cp_button   = null;

        private TText   op_name         = null;

        private int old_idx     = -1;

        //
        //
        //
        RoAllCopy(){
            setTitle("制御テーブル：全コピー");
            setSize(490,170);
            setResizable(false);
            setModal(true);
            getContentPane().setLayout(null);
            // 他基地参照機能    @20131021
            if(CZSystemDefine.REFERENCE_RUN == CZSystem.getRunLevel()){
                getContentPane().setBackground(CZSystemDefine.DEFAULT_REFERENCE_BACKGROUND_COL);
            }else{
                getContentPane().setBackground(CZSystemDefine.DEFAULT_BACKGROUND_COL);
            }

            addWindowListener(new WindowAdapter(){
                public void windowClosing(WindowEvent e){
                    winClose(e);
                }
            });

            JLabel  label   = null;

            label = new JLabel("コピー元",JLabel.CENTER);
            label.setBounds(20, 20, 100, 24);
            label.setLocale(new Locale("ja","JP"));
            label.setFont(new java.awt.Font("dialog", 0, 16));
            label.setBorder(new Flush3DBorder());
            label.setForeground(java.awt.Color.black);
            getContentPane().add(label);

            src_ro_name = new JLabel(" ",JLabel.CENTER);
            src_ro_name.setBounds(20, 54, 100, 24);
            src_ro_name.setLocale(new Locale("ja","JP"));
            src_ro_name.setFont(new java.awt.Font("dialog", 0, 16));
            src_ro_name.setBorder(new Flush3DBorder());
            src_ro_name.setForeground(java.awt.Color.black);
            getContentPane().add(src_ro_name);

            ///////////////////////////////////////////////////////////
            label = new JLabel("コピー先",JLabel.CENTER);
            label.setBounds(260, 20, 100, 24);
            label.setLocale(new Locale("ja","JP"));
            label.setFont(new java.awt.Font("dialog", 0, 16));
            label.setBorder(new Flush3DBorder());
            label.setForeground(java.awt.Color.black);
            getContentPane().add(label);

            dst_ro_name = new RoNo();
            dst_ro_name.setBounds(260, 54, 100, 24);
            getContentPane().add(dst_ro_name);

            ///////////////////////////////////////////////////////////
            label = new JLabel("設定者",JLabel.CENTER);
            label.setBounds(20, 110, 100, 24);
            label.setLocale(new Locale("ja","JP"));
            label.setFont(new java.awt.Font("dialog", 0, 16));
            label.setBorder(new Flush3DBorder());
            label.setForeground(java.awt.Color.black);
            getContentPane().add(label);

            // オペレータ名
            op_name = new TText();
            op_name.setBounds(120, 110, 140, 24);
            getContentPane().add(op_name);

            cp_button = new JButton();
            cp_button = new JButton("実  行");
            cp_button.setBounds(260, 110, 100, 24);
            cp_button.setLocale(new Locale("ja","JP"));
            cp_button.setFont(new java.awt.Font("dialog", 0, 18));
            cp_button.setBorder(new Flush3DBorder());
            cp_button.setForeground(java.awt.Color.black);
            cp_button.addActionListener(new SendButton());
            getContentPane().add(cp_button);

            // 終了
            JButton button = new JButton("終  了");
            button.setBounds(370, 110, 100, 24);
            button.setLocale(new Locale("ja","JP"));
            button.setFont(new java.awt.Font("dialog", 0, 18));
            button.setBorder(new Flush3DBorder());
            button.setForeground(java.awt.Color.black);
            button.addActionListener(new CancelButton());
            getContentPane().add(button);

        }


        //
        //
        //
        private void winClose(WindowEvent e){
            CZSystem.log("CZControlTableCp","RoAllCopy winClose() " + e);
            dst_ro_name.releaseHaita();
        }


        //
        //
        //
        public boolean setDefault(){

//@@            CZSystem.log("CZControlTableCp","RoAllCopy setDefault");
            op_name.setText("");

			String s = CZSystem.RoKetaChg(CZSystem.getRoName());	// 20050725 炉：表示桁数変更
            src_ro_name.setText(s);
//            src_ro_name.setText(CZSystem.getRoName());
            dst_ro_name.setDefault();
            return true;
        }


        //
        //
        //
        public boolean setSendStatus(){
            int idx = 0;
            idx = dst_ro_name.getSelectedIndex();
            if(0 > idx) return false;
            String sendOp = op_name.getText();
            if(1 > sendOp.length()) return false;
            String ro     = CZSystem.getRoName();
            String dst_ro = CZSystem.getRoName(idx);
            if( ro.equals(dst_ro) ) return false;
//@@            CZSystem.log("CZControlTableCp","RoAllCopy ro[" + ro + "]->[" + dst_ro + "]");
            return true;
        }


        //
        // エラーメッセージの表示
        //
        private boolean errorMsg(Object msg[]){
            JOptionPane.showMessageDialog(null,msg,
                        "制御テーブル：全コピー",
                        JOptionPane.ERROR_MESSAGE);
            return true;
        }


        //
        // 確認メッセージの表示
        //
        private boolean confirmDia(Object msg[]){

            int ans = JOptionPane.showConfirmDialog(null,msg,
                "制御テーブル：全コピー",
                JOptionPane.OK_CANCEL_OPTION,
                JOptionPane.WARNING_MESSAGE);
            if(0 == ans) return true;
            return false;
        }


        //
        //
        //
        public class RoNo extends JComboBox {

            RoNo(){
                super();

                try{
                    setName("JComboBox1");
                    setFont(new java.awt.Font("dialog", 0, 16));
                    Vector ro = CZSystem.getRoNameList();
                    if(null == ro){
                        CZSystem.exit(0,"Not Ro No");
                    }
                    for(int i = 0 ; ro.size() > i ; i++){
						String s = CZSystem.RoKetaChg((String)ro.elementAt(i));	// 20050725 炉：表示桁数変更
						addItem(s);
//                        addItem((String)ro.elementAt(i));
                    }
                    setForeground(java.awt.Color.black);
                    setBackground(java.awt.Color.lightGray);
					setFocusable(false);	/* 2007.08.22 */
                    addActionListener(new ChgRoNo());
//@@                    CZSystem.log("CZControlTableCp","RoAllCopy new RoNo()");
                }
                catch (Throwable e) {
                    CZSystem.handleException(e);
                }
            }


            //
            //
            //
            public void setDefault(){

                int idx = getSelectedIndex();
                old_idx = idx;      //@@@
//@@                CZSystem.log("CZControlTableCp","RoAllCopy setDefault() RoNo[" +
//@@                                idx + "][" + CZSystem.getRoName(idx) + "]" );
				String s = CZSystem.RoKetaChg(CZSystem.getRoName(idx));	// 20050725 炉：表示桁数変更

                if(getHaita(idx)){
                    cp_button.setEnabled(true);
                }
                else {
                    cp_button.setEnabled(false);
                    Object msg[] = {"全コピー",
                                    new String(s + "は"),
//                                    new String(CZSystem.getRoName(idx) + "は"),
                                "制御盤、他の端末で修正中です"};
                    errorMsg(msg);
                }
            }


            //
            // 画面消去時の排他開放
            //
            public void releaseHaita(){
                int idx = getSelectedIndex();
//@@                CZSystem.log("CZControlTableCp","RoAllCopy releaseHaita() 排他[" + idx + "]開放");
                putHaita(idx);
            }


            //
            //
            //
            class ChgRoNo implements ActionListener {
                public void actionPerformed(ActionEvent e){
                    RoNo obj = (RoNo)e.getSource();
                    if(-1 == old_idx){
//@@                        CZSystem.log("CZControlTableCp","RoAllCopy ChgRoNo() 排他１回目");
                    }
                    else {
                        putHaita(old_idx);
//@@                        CZSystem.log("CZControlTableCp","RoAllCopy ChgRoNo() 排他[" +
//@@                                        old_idx + "]->[" + obj.getSelectedIndex() + "]");
                    }
                    obj.setDefault();
                    old_idx = obj.getSelectedIndex();
                }
            }
        } // RoNo


        /***************************************************
         *
         *   実行ボタン
         *
         ***************************************************/
        class SendButton implements ActionListener {
            public void actionPerformed(ActionEvent ev){

                if(!setSendStatus()){
                    Object msg1[] = {"全コピー",
                                    "炉番、設定者を",
                                    "見直してください"};
                    errorMsg(msg1);
                    return;
                }

                //Send
                String sendOp = op_name.getText();
                int idx = dst_ro_name.getSelectedIndex();
                String dst_ro = CZSystem.getRoName(idx);

                Object msg2[] = {"コピーを開始します。下記の項目を確認してください！！",
                                "  1) コピー先の炉番は [" + dst_ro + "] ですか？"};
                if(!confirmDia(msg2)) return;
                if(!CZSystem.CZControlCopyRo(sendOp,dst_ro)){
                    dst_ro_name.setDefault();
                    Object msg[] = {"全コピー",
                                    "コピーが失敗しました",
                                    "管理者に問い合わせてください"};
                    errorMsg(msg);
                    return;
                }
                dst_ro_name.setDefault();
                return ;
            }
        }



        /***************************************************
         *
         *   終了ボタン
         *
         ***************************************************/
        class CancelButton implements ActionListener {
            public void actionPerformed(ActionEvent ev){
                setVisible(false);
                dst_ro_name.releaseHaita();
            }
        }
    } /* public class RoAllCopy extends JDialog */

    //////////////////////////////////////////////////////////////////////////////////////////////
    //
    //  グループコピー用画面
    //
    public class GroupCopy extends JDialog {

        private int groupe_no   = 0;
        private String  grupe_title = null;
        private JLabel  src_ro_name = null;
        private JLabel  src_grp_name    = null;
        private RoNo    dst_ro_name = null;
        private JButton cp_button   = null;
        private TText   op_name         = null;
        private int old_idx     = -1;

        //
        //
        //
        GroupCopy(){
            setTitle("制御テーブル：グループコピー");
            setSize(490,170);
            setResizable(false);
            setModal(true);
            getContentPane().setLayout(null);
            // 他基地参照機能    @20131021
            if(CZSystemDefine.REFERENCE_RUN == CZSystem.getRunLevel()){
                getContentPane().setBackground(CZSystemDefine.DEFAULT_REFERENCE_BACKGROUND_COL);
            }else{
                getContentPane().setBackground(CZSystemDefine.DEFAULT_BACKGROUND_COL);
            }

            addWindowListener(new WindowAdapter(){
                public void windowClosing(WindowEvent e){
                    winClose(e);
                }
            });

            JLabel  label   = null;

            label = new JLabel("コピー元",JLabel.CENTER);
            label.setBounds(20, 20, 100, 24);
            label.setLocale(new Locale("ja","JP"));
            label.setFont(new java.awt.Font("dialog", 0, 16));
            label.setBorder(new Flush3DBorder());
            label.setForeground(java.awt.Color.black);
            getContentPane().add(label);

            src_ro_name = new JLabel(" ",JLabel.CENTER);
            src_ro_name.setBounds(20, 54, 100, 24);
            src_ro_name.setLocale(new Locale("ja","JP"));
            src_ro_name.setFont(new java.awt.Font("dialog", 0, 16));
            src_ro_name.setBorder(new Flush3DBorder());
            src_ro_name.setForeground(java.awt.Color.black);
            getContentPane().add(src_ro_name);

            src_grp_name = new JLabel(" ",JLabel.CENTER);
            src_grp_name.setBounds(140, 54, 100, 24);
            src_grp_name.setLocale(new Locale("ja","JP"));
            src_grp_name.setFont(new java.awt.Font("dialog", 0, 16));
            src_grp_name.setBorder(new Flush3DBorder());
            src_grp_name.setForeground(java.awt.Color.black);
            getContentPane().add(src_grp_name);

            ///////////////////////////////////////////////////////////
            label = new JLabel("コピー先",JLabel.CENTER);
            label.setBounds(260, 20, 100, 24);
            label.setLocale(new Locale("ja","JP"));
            label.setFont(new java.awt.Font("dialog", 0, 16));
            label.setBorder(new Flush3DBorder());
            label.setForeground(java.awt.Color.black);
            getContentPane().add(label);

            dst_ro_name = new RoNo();
            dst_ro_name.setBounds(260, 54, 100, 24);
            getContentPane().add(dst_ro_name);

            ///////////////////////////////////////////////////////////
            label = new JLabel("設定者",JLabel.CENTER);
            label.setBounds(20, 110, 100, 24);
            label.setLocale(new Locale("ja","JP"));
            label.setFont(new java.awt.Font("dialog", 0, 16));
            label.setBorder(new Flush3DBorder());
            label.setForeground(java.awt.Color.black);
            getContentPane().add(label);

            // オペレータ名
            op_name = new TText();
            op_name.setBounds(120, 110, 140, 24);
            getContentPane().add(op_name);

            cp_button = new JButton();
            cp_button = new JButton("実  行");
            cp_button.setBounds(260, 110, 100, 24);
            cp_button.setLocale(new Locale("ja","JP"));
            cp_button.setFont(new java.awt.Font("dialog", 0, 18));
            cp_button.setBorder(new Flush3DBorder());
            cp_button.setForeground(java.awt.Color.black);
            cp_button.addActionListener(new SendButton());
            getContentPane().add(cp_button);

            // 終了
            JButton button = new JButton("終  了");
            button.setBounds(370, 110, 100, 24);
            button.setLocale(new Locale("ja","JP"));
            button.setFont(new java.awt.Font("dialog", 0, 18));
            button.setBorder(new Flush3DBorder());
            button.setForeground(java.awt.Color.black);
            button.addActionListener(new CancelButton());
            getContentPane().add(button);
        }


        //
        //
        //
        private void winClose(WindowEvent e){
            CZSystem.log("CZControlTableCp","GroupCopy winClose() " + e);
            dst_ro_name.releaseHaita();
        }


        //
        //
        //
        public boolean setDefault(int grp_no,String grp){

//@@            CZSystem.log("CZControlTableCp","GroupCopy grp_no[" + grp_no +"] grp[" + grp + "]");
            op_name.setText("");
            groupe_no   = grp_no;
            grupe_title = grp;
            
            String s = CZSystem.RoKetaChg(CZSystem.getRoName());	// 20050725 炉：表示桁数変更
            src_ro_name.setText(s);
//            src_ro_name.setText(CZSystem.getRoName());
            src_grp_name.setText(grupe_title);
            dst_ro_name.setDefault();
            return true;
        }


        //
        //
        //
        public boolean setSendStatus(){

            int idx = 0;
            idx = dst_ro_name.getSelectedIndex();
            if(0 > idx) return false;
            String sendOp = op_name.getText();
            if(1 > sendOp.length()) return false;
            String ro     = CZSystem.getRoName();
            String dst_ro = CZSystem.getRoName(idx);
            if( ro.equals(dst_ro) ) return false;
//@@            CZSystem.log("CZControlTableCp","GroupCopy setSendStatus() grp_no[" +
//@@                                groupe_no + "] ro [" + ro + "]->[" + dst_ro + "]");
            return true;
        }

        //
        // エラーメッセージの表示
        //
        private boolean errorMsg(Object msg[]){
            JOptionPane.showMessageDialog(null,msg,
                            "制御テーブル：グループコピー",
                            JOptionPane.ERROR_MESSAGE);
            return true;
        }


        //
        // 確認メッセージの表示
        //
        private boolean confirmDia(Object msg[]){

            int ans = JOptionPane.showConfirmDialog(null,msg,
                    "制御テーブル：グループコピー",
                    JOptionPane.OK_CANCEL_OPTION,
                    JOptionPane.WARNING_MESSAGE);
            if(0 == ans) return true;
            return false;
        }



        //
        //
        //
        public class RoNo extends JComboBox {

            RoNo(){
                super();
                try{
                    setName("JComboBox1");
                    setFont(new java.awt.Font("dialog", 0, 16));
                    Vector ro = CZSystem.getRoNameList();
                    if(null == ro){
                        CZSystem.exit(0,"Not Ro No");
                    }
                    for(int i = 0 ; ro.size() > i ; i++){
						String s = CZSystem.RoKetaChg((String)ro.elementAt(i));	// 20050725 炉：表示桁数変更
                        addItem(s);
//                        addItem((String)ro.elementAt(i));
                    }
                    setForeground(java.awt.Color.black);
                    setBackground(java.awt.Color.lightGray);
					setFocusable(false);	/* 2007.08.22 */
                    addActionListener(new ChgRoNo());
//@@                    CZSystem.log("CZControlTableCp","GroupCopy new RoNo()");
                }
                catch (Throwable e) {
                    CZSystem.handleException(e);
                }
            }


            //
            //
            //
            public void setDefault(){

                int idx = getSelectedIndex();
                old_idx = idx;      //@@@
//@@                CZSystem.log("CZControlTableCp","GroupCopy setDefault() RoNo[" +
//@@                                        idx + "][" + CZSystem.getRoName(idx) + "]" );
				String s = CZSystem.RoKetaChg(CZSystem.getRoName(idx));	// 20050725 炉：表示桁数変更

                if(getHaita(idx)){
                    cp_button.setEnabled(true);
                }
                else {
                    cp_button.setEnabled(false);
                    Object msg[] = {"グループコピー",
                                    new String(s + "は"),
//                                    new String(CZSystem.getRoName(idx) + "は"),
                                "制御盤、他の端末で修正中です"};
                    errorMsg(msg);
                }
            }


            //
            // 画面消去時の排他開放
            //
            public void releaseHaita(){
                int idx = getSelectedIndex();
//@@                CZSystem.log("CZControlTableCp","GroupCopy releaseHaita() 排他[" + idx + "]開放");
                putHaita(idx);
            }


            //
            //
            //
            class ChgRoNo implements ActionListener {
                public void actionPerformed(ActionEvent e){

                    RoNo obj = (RoNo)e.getSource();
                    if(-1 == old_idx){
//@@                        CZSystem.log("CZControlTableCp","GroupCopy ChgRoNo 排他１回目");
                    }
                    else {
                        putHaita(old_idx);
//@@                        CZSystem.log("CZControlTableCp","GroupCopy ChgRoNo 排他[" +
//@@                                        old_idx + "]->[" + obj.getSelectedIndex() + "]");
                    }
                    obj.setDefault();
                    old_idx = obj.getSelectedIndex();
                }
            }
        } // RoNo


        /***************************************************
         *
         *   実行ボタン
         *
         ***************************************************/
        class SendButton implements ActionListener {
            public void actionPerformed(ActionEvent ev){

                if(!setSendStatus()){
                    Object msg1[] = {"グループコピー",
                                    "炉番、設定者を",
                                    "見直してください"};
                    errorMsg(msg1);
                    return;
                }
                //Send
                String sendOp = op_name.getText();
                int idx = dst_ro_name.getSelectedIndex();
                String dst_ro = CZSystem.getRoName(idx);
                Object msg2[] = {"コピーを開始します。下記の項目を確認してください！！",
                                "  1) グループコピー   [" + grupe_title + "] でいいですか？",
                                "  2) コピー先の炉番は [" + dst_ro + "] ですか？"};
                if(!confirmDia(msg2)) return;
                if(!CZSystem.CZControlCopyGroup(sendOp,dst_ro,groupe_no)){
                    dst_ro_name.setDefault();
                    Object msg[] = {"グループコピー",
                                    "コピーが失敗しました",
                                    "管理者に問い合わせてください"};
                    errorMsg(msg);
                    return;
                }
                dst_ro_name.setDefault();
                return ;
            }
        }


        /***************************************************
         *
         *   終了ボタン
         *
         ***************************************************/
        class CancelButton implements ActionListener {
            public void actionPerformed(ActionEvent ev){
                setVisible(false);
                dst_ro_name.releaseHaita();
            }
        }
    } /* public class GroupCopy extends JDialog */

    //////////////////////////////////////////////////////////////////////////////////////////////
    //
    //  レシピコピー用画面
    //
    //////////////////////////////////////////////////////////////////////////////////////////////
    public class RecipeCopy extends JDialog {

        private int send_recipie_no = 0;

        private int groupe_no   = 0;
        private int recipie_no  = 0;
        private String  grupe_title = null;
        private String  recipie_title   = null;

        private JLabel  src_ro_name = null;
        private JLabel  src_grp_name    = null;
        private JLabel  src_rec_no  = null;
        private JLabel  src_grp_title   = null;

        private RoNo    dst_ro_name = null;
        private GroupeTable g_table = null;

        private JButton cp_button   = null;

        private Vector  table_title = null;

        private TText   op_name         = null;

        private int old_idx     = -1;

        //
        //
        //
        RecipeCopy(){
            setTitle("制御テーブル：レシピーコピー");
            setSize(710,480);
            setResizable(false);
            setModal(true);
            getContentPane().setLayout(null);
            // 他基地参照機能    @20131021
            if(CZSystemDefine.REFERENCE_RUN == CZSystem.getRunLevel()){
                getContentPane().setBackground(CZSystemDefine.DEFAULT_REFERENCE_BACKGROUND_COL);
            }else{
                getContentPane().setBackground(CZSystemDefine.DEFAULT_BACKGROUND_COL);
            }

            addWindowListener(new WindowAdapter(){
                public void windowClosing(WindowEvent e){
                    winClose(e);
                }
            });

            JLabel  label   = null;
            label = new JLabel("コピー元",JLabel.CENTER);
            label.setBounds(20, 20, 100, 24);
            label.setLocale(new Locale("ja","JP"));
            label.setFont(new java.awt.Font("dialog", 0, 16));
            label.setBorder(new Flush3DBorder());
            label.setForeground(java.awt.Color.black);
            getContentPane().add(label);

            src_ro_name = new JLabel(" ",JLabel.CENTER);
            src_ro_name.setBounds(20, 54, 100, 24);
            src_ro_name.setLocale(new Locale("ja","JP"));
            src_ro_name.setFont(new java.awt.Font("dialog", 0, 16));
            src_ro_name.setBorder(new Flush3DBorder());
            src_ro_name.setForeground(java.awt.Color.black);
            getContentPane().add(src_ro_name);

            src_grp_name = new JLabel(" ",JLabel.CENTER);
            src_grp_name.setBounds(140, 54, 100, 24);
            src_grp_name.setLocale(new Locale("ja","JP"));
            src_grp_name.setFont(new java.awt.Font("dialog", 0, 16));
            src_grp_name.setBorder(new Flush3DBorder());
            src_grp_name.setForeground(java.awt.Color.black);
            getContentPane().add(src_grp_name);

            src_rec_no = new JLabel(" ",JLabel.CENTER);
            src_rec_no.setBounds(260, 54, 100, 24);
            src_rec_no.setLocale(new Locale("ja","JP"));
            src_rec_no.setFont(new java.awt.Font("dialog", 0, 16));
            src_rec_no.setBorder(new Flush3DBorder());
            src_rec_no.setForeground(java.awt.Color.black);
            getContentPane().add(src_rec_no);

            src_grp_title = new JLabel(" ",JLabel.CENTER);
            src_grp_title.setBounds(140, 88, 540, 24);
            src_grp_title.setLocale(new Locale("ja","JP"));
            src_grp_title.setFont(new java.awt.Font("dialog", 0, 16));
            src_grp_title.setBorder(new Flush3DBorder());
            src_grp_title.setForeground(java.awt.Color.black);
            getContentPane().add(src_grp_title);

            label = new JLabel("コピー先",JLabel.CENTER);
            label.setBounds(20, 150, 100, 24);
            label.setLocale(new Locale("ja","JP"));
            label.setFont(new java.awt.Font("dialog", 0, 16));
            label.setBorder(new Flush3DBorder());
            label.setForeground(java.awt.Color.black);
            getContentPane().add(label);

            dst_ro_name = new RoNo();
            dst_ro_name.setBounds(20, 184, 100, 24);
            getContentPane().add(dst_ro_name);

            JScrollPane panel = null;
            g_table = new GroupeTable();
            panel = new JScrollPane(g_table);
            panel.setBounds(140, 184, 540, 200);
            getContentPane().add(panel);

            label = new JLabel("設定者",JLabel.CENTER);
            label.setBounds(20, 410, 100, 24);
            label.setLocale(new Locale("ja","JP"));
            label.setFont(new java.awt.Font("dialog", 0, 16));
            label.setBorder(new Flush3DBorder());
            label.setForeground(java.awt.Color.black);
            getContentPane().add(label);

            // オペレータ名
            op_name = new TText();
            op_name.setBounds(120, 410, 140, 24);
            getContentPane().add(op_name);

            cp_button = new JButton();
            cp_button = new JButton("実  行");
            cp_button.setBounds(260, 410, 100, 24);
            cp_button.setLocale(new Locale("ja","JP"));
            cp_button.setFont(new java.awt.Font("dialog", 0, 18));
            cp_button.setBorder(new Flush3DBorder());
            cp_button.setForeground(java.awt.Color.black);
            cp_button.addActionListener(new SendButton());
            getContentPane().add(cp_button);

            // 終了
            JButton button = new JButton("終  了");
            button.setBounds(580, 410, 100, 24);
            button.setLocale(new Locale("ja","JP"));
            button.setFont(new java.awt.Font("dialog", 0, 18));
            button.setBorder(new Flush3DBorder());
            button.setForeground(java.awt.Color.black);
            button.addActionListener(new CancelButton());
            getContentPane().add(button);
        }

        //
        //
        //
        private void winClose(WindowEvent e){
            CZSystem.log("CZControlTableCp","RecipeCopy winClose() " + e);
            dst_ro_name.releaseHaita();
        }


        //
        //
        //
        public boolean setDefault(int grp_no,String grp,int rec_no,String title){

//@@            CZSystem.log("CZControlTableCp","RecipeCopy grp_no[" + grp_no +"] grp[" + grp +
//@@                   "] rec_no[" + rec_no + "] title[" + title +"]");

            op_name.setText("");
            groupe_no   = grp_no;
            recipie_no  = rec_no;
            grupe_title = grp;
            recipie_title   = title;
            String s = CZSystem.RoKetaChg(CZSystem.getRoName());	// 20050725 炉：表示桁数変更
            src_ro_name.setText(s);
//            src_ro_name.setText(CZSystem.getRoName());
            src_grp_name.setText(grupe_title);
            src_rec_no.setText(new String(" " + recipie_no + " "));
            src_grp_title.setText(recipie_title);
            dst_ro_name.setDefault();
            return true;
        }


        //
        //
        //
        public boolean setSendStatus(){
            int idx = 0;
            idx = dst_ro_name.getSelectedIndex();
            if(0 > idx) return false;
            String sendOp = op_name.getText();
            if(1 > sendOp.length()) return false;
            send_recipie_no = g_table.getSelectedRow() + 1;
            if(1 > send_recipie_no) return false;
            String ro     = CZSystem.getRoName();
            String dst_ro = CZSystem.getRoName(idx);
            if( ro.equals(dst_ro) && (recipie_no == send_recipie_no)) return false;

//@@            CZSystem.log("CZControlTableCp","RecipeCopy grp_no[" + groupe_no  + "]");
//@@            CZSystem.log("CZControlTableCp","RecipeCopy ro    [" + ro + "]->[" + dst_ro + "]");
//@@            CZSystem.log("CZControlTableCp","RecipeCopy rec_no[" + recipie_no + "]->[" + send_recipie_no + "]");

            return true;
        }


        //
        // メッセージの表示
        //
        private boolean errorMsg(Object msg[]){
            JOptionPane.showMessageDialog(null,msg,
                    "制御テーブル：レシピーコピー",
                    JOptionPane.ERROR_MESSAGE);
            return true;
        }


        //
        //
        //
        public class RoNo extends JComboBox {

            RoNo(){
                super();
                try{
                    setName("JComboBox1");
                    setFont(new java.awt.Font("dialog", 0, 16));
                    Vector ro = CZSystem.getRoNameList();
                    if(null == ro){
                        CZSystem.exit(0,"Not Ro No");
                    }
                    for(int i = 0 ; ro.size() > i ; i++){
						String s = CZSystem.RoKetaChg((String)ro.elementAt(i));	// 20050725 炉：表示桁数変更
                        addItem(s);
//                        addItem((String)ro.elementAt(i));
                    }
                    setForeground(java.awt.Color.black);
                    setBackground(java.awt.Color.lightGray);
					setFocusable(false);	/* 2007.08.22 */
                    addActionListener(new ChgRoNo());
//@@                    CZSystem.log("CZControlTableCp","RecipeCopy new RoNo()");
                }
                catch (Throwable e) {
                    CZSystem.handleException(e);
                }
            }


            //
            //
            //
            public void setDefault(){

                int idx = getSelectedIndex();
                old_idx = idx;      //@@@
//@@                CZSystem.log("CZControlTableCp","RecipeCopy setDefault() RoNo[" +
//@@                                idx + "][" + CZSystem.getRoName(idx) + "]" );
				String s = CZSystem.RoKetaChg(CZSystem.getRoName(idx));	// 20050725 炉：表示桁数変更

                table_title = null;
                table_title  = CZSystem.getCtTitle(idx);
                if(null != table_title){
                    g_table.setData(groupe_no,recipie_no);
                }
                if(getHaita(idx)){
                    cp_button.setEnabled(true);
                }
                else {
                    cp_button.setEnabled(false);
                    Object msg[] = {"レシピーコピー",
                        new String(s + "は"),
//                        new String(CZSystem.getRoName(idx) + "は"),
                        "制御盤、他の端末で修正中です"};
                    errorMsg(msg);
                }
            }


            //
            // 画面消去時の排他開放
            //
            public void releaseHaita(){
                int idx = getSelectedIndex();
//@@                CZSystem.log("CZControlTableCp","RecipeCopy releaseHaita() 排他[" + idx + "]開放");
                putHaita(idx);
            }


            //
            //
            //
            class ChgRoNo implements ActionListener {
                public void actionPerformed(ActionEvent e){

                    RoNo obj = (RoNo)e.getSource();
                    if(-1 == old_idx){
//@@                        CZSystem.log("CZControlTableCp","RecipeCopy ChgRoNo 排他１回目");
                    }
                    else {
                        putHaita(old_idx);
//@@                        CZSystem.log("CZControlTableCp","RecipeCopy ChgRoNo 排他[" +
//@@                                        old_idx + "]->[" + obj.getSelectedIndex() + "]");
                    }
                    obj.setDefault();
                    old_idx = obj.getSelectedIndex();
                }
            }

        }


        /***************************************************
         *
         *       グループ別のレシピテーブル一覧
         *
         ***************************************************/
        class GroupeTable extends JTable {

            private GrTblMdl model = null;

            GroupeTable(){
                super();
                try{
                    setName("GroupeTable");
                    setBounds(0, 0, 200, 200);
                    setAutoCreateColumnsFromModel(true);
                    setSelectionMode(ListSelectionModel.SINGLE_SELECTION);
                    setLocale(new Locale("ja","JP"));
                    setFont(new java.awt.Font("dialog", 0, 12));
                    setRowHeight(17);
                    model = new GrTblMdl();
                    setModel(model);
                    DefaultTableColumnModel cmdl = (DefaultTableColumnModel)getColumnModel();
                    TableColumn  colum = null;
                    // レシピーNo
                    colum = cmdl.getColumn(0);
                    colum.setMaxWidth(40);
                    colum.setMinWidth(40);
                    colum.setWidth(40);
                    // タイトル
                    colum = cmdl.getColumn(1);
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
//@@                CZSystem.log("CZControlTableCp","GroupeTable valueChanged [" +
//@@                            getSelectedRow() + "][" + getSelectedColumn() + "]");
            }

            //
            //
            //
            public void setData(int gr,int tbl){

//@@                CZSystem.log("CZControlTableCp","GroupeTable setData() [" + gr + "][" + tbl + "]");
                CZSystemCtTitle title   = null;
                String          empty   = new String("");
                for(int i = 0 ; i < 999 ; i++){
                    g_table.setValueAt(empty,i,1);
                }
                for(int i = 0 ; i < table_title.size() ; i++){
                    title = (CZSystemCtTitle)table_title.elementAt(i);
                    if(gr == title.g_no){
                        g_table.setValueAt(title.title,title.r_no-1,1);
                    }
                }
                setRowSelectionInterval(tbl-1,tbl-1);
                Rectangle cellRect = getCellRect(tbl-1,0,false);
                if(cellRect != null){
                    scrollRectToVisible(cellRect);
                }
                repaint();
            }
        }

        /***************************************************
         *
         *       制御テーブルクラス：モデル
         *
         ***************************************************/
        public class GrTblMdl extends AbstractTableModel {

            final   int TBL_COL     = 2;
            private int TBL_ROW     = 999;
            final String[] names = {" # " ,"タイトル"};
            private Object  data[][];

            //
            GrTblMdl(){
                super();
                data = new Object[TBL_ROW][TBL_COL];
                for(int i = 0 ; i < TBL_ROW ; i++){
                    data[i][0] = new Integer(i+1);
                    data[i][1] = new String("");
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

        /***************************************************
         *
         *   実行ボタン
         *
         ***************************************************/
        class SendButton implements ActionListener {
            public void actionPerformed(ActionEvent ev){

                if(!setSendStatus()){
                    Object msg[] = {"レシピーコピー",
                                    "レシピー、設定者を",
                                    "見直してください"};
                    errorMsg(msg);
                    return;
                 }
                //Send
                String sendOp = op_name.getText();
                int idx = dst_ro_name.getSelectedIndex();
                String dst_ro = CZSystem.getRoName(idx);
                if(!CZSystem.CZControlCopyRecipe(sendOp,dst_ro,groupe_no,recipie_no,send_recipie_no)){
                    dst_ro_name.setDefault();
                    Object msg[] = {"レシピーコピー",
                                    "コピーが失敗しました",
                                    "管理者に問い合わせてください"};
                    errorMsg(msg);
                    return;
                }
                dst_ro_name.setDefault();
                return ;
            }
        }



        /***************************************************
         *
         *   終了ボタン
         *
         ***************************************************/
        class CancelButton implements ActionListener {
            public void actionPerformed(ActionEvent ev){
                setVisible(false);
                dst_ro_name.releaseHaita();
            }
        }
    } /* public class RecipeCopy extends JDialog */


    ////////////////////////////////////////////////////////
    //
    //  テーブルコピー用画面
    //
    ////////////////////////////////////////////////////////
    public class TableCopy extends JDialog {

        private int send_recipie_no = 0;
        private int groupe_no   = 0;
        private int recipie_no  = 0;
/// 2011.04.12 Y.K start
//        private int table_no    = 0;
//		private SelectTbTable t_table = null;
        private JScrollPane panelFromTbl = null;
/// 2011.04.12 Y.K end
        private String  grupe_title = null;
        private String  recipie_title   = null;
        private String  src_table_title = null;
        private JLabel  src_ro_name = null;
        private JLabel  src_grp_name    = null;
        private JLabel  src_rec_no  = null;
        private JLabel  src_grp_title   = null;
        private JLabel  src_tbl_no  = null;
        private JLabel  src_tbl_title   = null;
        private RoNo    dst_ro_name = null;
        private GroupeTable g_table = null;
        private JButton cp_button   = null;
        private Vector  table_title = null;
        private TText   op_name         = null;
        private int old_idx     = -1;

        //
        //
        //
        TableCopy(){
            setTitle("制御テーブル：テーブルコピー");
            setSize(710,510 + 84);
            setResizable(false);
            setModal(true);
            getContentPane().setLayout(null);
            // 他基地参照機能    @20131021
            if(CZSystemDefine.REFERENCE_RUN == CZSystem.getRunLevel()){
                getContentPane().setBackground(CZSystemDefine.DEFAULT_REFERENCE_BACKGROUND_COL);
            }else{
                getContentPane().setBackground(CZSystemDefine.DEFAULT_BACKGROUND_COL);
            }

            addWindowListener(new WindowAdapter(){
                public void windowClosing(WindowEvent e){
                    winClose(e);
                }
            });

            JLabel  label   = null;

            label = new JLabel("コピー元",JLabel.CENTER);
            label.setBounds(20, 20, 100, 24);
            label.setLocale(new Locale("ja","JP"));
            label.setFont(new java.awt.Font("dialog", 0, 16));
            label.setBorder(new Flush3DBorder());
            label.setForeground(java.awt.Color.black);
            getContentPane().add(label);

            src_ro_name = new JLabel(" ",JLabel.CENTER);
            src_ro_name.setBounds(20, 54, 100, 24);
            src_ro_name.setLocale(new Locale("ja","JP"));
            src_ro_name.setFont(new java.awt.Font("dialog", 0, 16));
            src_ro_name.setBorder(new Flush3DBorder());
            src_ro_name.setForeground(java.awt.Color.black);
            getContentPane().add(src_ro_name);

            src_grp_name = new JLabel(" ",JLabel.CENTER);
            src_grp_name.setBounds(140, 54, 100, 24);
            src_grp_name.setLocale(new Locale("ja","JP"));
            src_grp_name.setFont(new java.awt.Font("dialog", 0, 16));
            src_grp_name.setBorder(new Flush3DBorder());
            src_grp_name.setForeground(java.awt.Color.black);
            getContentPane().add(src_grp_name);

            src_rec_no = new JLabel(" ",JLabel.CENTER);
            src_rec_no.setBounds(260, 54, 100, 24);
            src_rec_no.setLocale(new Locale("ja","JP"));
            src_rec_no.setFont(new java.awt.Font("dialog", 0, 16));
            src_rec_no.setBorder(new Flush3DBorder());
            src_rec_no.setForeground(java.awt.Color.black);
            getContentPane().add(src_rec_no);

            src_grp_title = new JLabel(" ",JLabel.CENTER);
            src_grp_title.setBounds(140, 88, 540, 24);
            src_grp_title.setLocale(new Locale("ja","JP"));
            src_grp_title.setFont(new java.awt.Font("dialog", 0, 16));
            src_grp_title.setBorder(new Flush3DBorder());
            src_grp_title.setForeground(java.awt.Color.black);
            getContentPane().add(src_grp_title);

/// 2011.04.12 Y.K start
//            t_table = new SelectTbTable();
//            panelFromTbl = new JScrollPane(t_table);
            panelFromTbl = new JScrollPane();
            panelFromTbl.setBounds(140, 122, 540, 110);
            getContentPane().add(panelFromTbl);

//            src_tbl_no = new JLabel(" ",JLabel.CENTER);
//            src_tbl_no.setBounds(140, 122, 100, 24);
//            src_tbl_no.setLocale(new Locale("ja","JP"));
//            src_tbl_no.setFont(new java.awt.Font("dialog", 0, 16));
//            src_tbl_no.setBorder(new Flush3DBorder());
//            src_tbl_no.setForeground(java.awt.Color.black);
//            getContentPane().add(src_tbl_no);

//            src_tbl_title = new JLabel(" ",JLabel.CENTER);
//            src_tbl_title.setBounds(260, 122, 420, 24);
//            src_tbl_title.setLocale(new Locale("ja","JP"));
//            src_tbl_title.setFont(new java.awt.Font("dialog", 0, 16));
//            src_tbl_title.setBorder(new Flush3DBorder());
//            src_tbl_title.setForeground(java.awt.Color.black);
//            getContentPane().add(src_tbl_title);
/// 2011.04.12 Y.K end

            label = new JLabel("コピー先",JLabel.CENTER);
            label.setBounds(20, 180 + 70, 100, 24);
            label.setLocale(new Locale("ja","JP"));
            label.setFont(new java.awt.Font("dialog", 0, 16));
            label.setBorder(new Flush3DBorder());
            label.setForeground(java.awt.Color.black);
            getContentPane().add(label);

            dst_ro_name = new RoNo();
            dst_ro_name.setBounds(20, 214 + 70, 100, 24);
            getContentPane().add(dst_ro_name);

            JScrollPane panel = null;
            g_table = new GroupeTable();
            panel = new JScrollPane(g_table);
            panel.setBounds(140, 214 + 70, 540, 200);
            getContentPane().add(panel);

            label = new JLabel("設定者",JLabel.CENTER);
            label.setBounds(20, 440 + 70, 100, 24);
            label.setLocale(new Locale("ja","JP"));
            label.setFont(new java.awt.Font("dialog", 0, 16));
            label.setBorder(new Flush3DBorder());
            label.setForeground(java.awt.Color.black);
            getContentPane().add(label);

            // オペレータ名
            op_name = new TText();
            op_name.setBounds(120, 440 + 70, 140, 24);
            getContentPane().add(op_name);

            cp_button = new JButton();
            cp_button = new JButton("実  行");
            cp_button.setBounds(260, 440 + 70, 100, 24);
            cp_button.setLocale(new Locale("ja","JP"));
            cp_button.setFont(new java.awt.Font("dialog", 0, 18));
            cp_button.setBorder(new Flush3DBorder());
            cp_button.setForeground(java.awt.Color.black);
            cp_button.addActionListener(new SendButton());
            getContentPane().add(cp_button);

            // 終了
            JButton button = new JButton("終  了");
            button.setBounds(580, 440 + 70, 100, 24);
            button.setLocale(new Locale("ja","JP"));
            button.setFont(new java.awt.Font("dialog", 0, 18));
            button.setBorder(new Flush3DBorder());
            button.setForeground(java.awt.Color.black);
            button.addActionListener(new CancelButton());
            getContentPane().add(button);
        }


        //
        //
        //
        private void winClose(WindowEvent e){
            CZSystem.log("CZControlTableCp","TableCopy winClose() " + e);
            dst_ro_name.releaseHaita();
        }
//// 2011.04.12 y.k start
	    // バッチ情報を作成する。
	    public boolean setBtCondition(int[] TblNo, String[] TblName){

	        removeBtCondition();

	        SelectTbTable t = new SelectTbTable(TblNo,TblName);
	        JTableHeader tabHead = t.getTableHeader();
	        tabHead.setReorderingAllowed(false);
	        panelFromTbl.setViewportView(t);

	        return true;
	    }

	    //
	    // バッチ情報を削除する。
	    public boolean removeBtCondition(){

	        JViewport v;
	        v =  panelFromTbl.getViewport();
	        if(null != v.getView()) v.remove(v.getView());

	        return true;
	    }
/// 2011.04.12 y.k end

        //
        //
        //
        public boolean setDefault(int grp_no,String grp,int rec_no,String title,
                          int[] tbl_no,String[] t_title){
// 2011.04.12 Y.K                          int tbl_no,String t_title){

//@@            CZSystem.log("CZControlTableCP","TableCopy grp_no[" + grp_no +"] grp[" + grp +
//@@                       "] rec_no[" + rec_no + "] title[" + title +"] tbl_no[" + tbl_no + "] t_title[" + t_title + "]");

            op_name.setText("");

            groupe_no   = grp_no;
            recipie_no  = rec_no;
////2011.04.12 Y.K start
//            table_no    = tbl_no;
////2011.04.12 Y.K end
            grupe_title = grp;
            recipie_title   = title;
////2011.04.12 Y.K start
//            src_table_title = t_title;
////2011.04.12 Y.K end

			String s = CZSystem.RoKetaChg(CZSystem.getRoName());	// 20050725 炉：表示桁数変更
            src_ro_name.setText(s);
//            src_ro_name.setText(CZSystem.getRoName());
            src_grp_name.setText(grupe_title);
            src_rec_no.setText(new String(" " + recipie_no + " "));
            src_grp_title.setText(recipie_title);
////2011.04.12 Y.K start
			removeBtCondition();
			setBtCondition(tbl_no,t_title);
//            src_tbl_no.setText(new String(" " + table_no + " "));
//            src_tbl_title.setText(src_table_title);
////2011.04.12 Y.K End

            dst_ro_name.setDefault();
            return true;
        }


        //
        //
        //
        public boolean setSendStatus(){
            int idx = 0;
            idx = dst_ro_name.getSelectedIndex();
            if(0 > idx) return false;
            String sendOp = op_name.getText();
            if(1 > sendOp.length()) return false;
            send_recipie_no = g_table.getSelectedRow() + 1;
            if(1 > send_recipie_no) return false;
            String ro     = CZSystem.getRoName();
            String dst_ro = CZSystem.getRoName(idx);
            if( ro.equals(dst_ro) && (recipie_no == send_recipie_no)) return false;
//@@            CZSystem.log("CZControlTableCP","TableCopy grp_no[" + groupe_no  + "]");
//@@            CZSystem.log("CZControlTableCP","TableCopy ro    [" + ro + "]->[" + dst_ro + "]");
//@@            CZSystem.log("CZControlTableCP","TableCopy rec_no[" + recipie_no + "]->[" + send_recipie_no + "]");
            return true;
        }


        //
        // メッセージの表示
        //
        private boolean errorMsg(Object msg[]){
            JOptionPane.showMessageDialog(null,msg,
                        "制御テーブル：テーブルコピー",
                        JOptionPane.ERROR_MESSAGE);
            return true;
        }


        //
        //
        //
        public class RoNo extends JComboBox {

            // ---------- コンストラクタ -------------------
            //
            RoNo(){
                super();

                try{
                    setName("JComboBox1");
                    setFont(new java.awt.Font("dialog", 0, 16));

                    Vector ro = CZSystem.getRoNameList();
                    if(null == ro){
                        CZSystem.exit(0,"Not Ro No");
                    }

                    for(int i = 0 ; ro.size() > i ; i++){
						String s = CZSystem.RoKetaChg((String)ro.elementAt(i));	// 20050725 炉：表示桁数変更
                        addItem(s);
//                        addItem((String)ro.elementAt(i));
                    }

                    setForeground(java.awt.Color.black);
                    setBackground(java.awt.Color.lightGray);
					setFocusable(false);	/* 2007.08.22 */
                    addActionListener(new ChgRoNo());
//@@                    CZSystem.log("CZControlTableCP","TableCopy new RoNo()");
                }
                catch (Throwable e) {
                    CZSystem.handleException(e);
                }
            }


            //
            //
            //
            public void setDefault(){

                int idx = getSelectedIndex();
                old_idx = idx;      //@@@

//@@                CZSystem.log("CZControlTableCP","TableCopy setDefault() RoNo[" +
//@@                                idx + "][" + CZSystem.getRoName(idx) + "]" );
				String s = CZSystem.RoKetaChg(CZSystem.getRoName(idx));	// 20050725 炉：表示桁数変更

                table_title = null;
                table_title  = CZSystem.getCtTitle(idx);
                if(null != table_title){
                    g_table.setData(groupe_no,recipie_no);
                }
                if(getHaita(idx)){
                    cp_button.setEnabled(true);
                }
                else {
                    cp_button.setEnabled(false);
                    Object msg[] = {"テーブルコピー",
                                new String(s + "は"),
//                                new String(CZSystem.getRoName(idx) + "は"),
                                    "制御盤、他の端末で修正中です"};
                    errorMsg(msg);
                }
            }


            //
            // 画面消去時の排他開放
            //
            public void releaseHaita(){
                int idx = getSelectedIndex();
//@@                CZSystem.log("CZControlTableCP","TableCopy releaseHaita() 排他[" + idx + "]開放");
                putHaita(idx);
            }


            //
            //
            //
            class ChgRoNo implements ActionListener {
                public void actionPerformed(ActionEvent e){

                    RoNo obj = (RoNo)e.getSource();

                    if(-1 == old_idx){
//@@                        CZSystem.log("CZControlTableCP","TableCopy ChgRoNo() 排他１回目");
                    }
                    else {
                        putHaita(old_idx);
//@@                        CZSystem.log("CZControlTableCP","TableCopy ChgRoNo() 排他[" +
//@@                                        old_idx + "]->[" + obj.getSelectedIndex() + "]");
                    }
                    obj.setDefault();
                    old_idx = obj.getSelectedIndex();
                }
            }
        }

////2011.04.12 Y.K Start
        /***************************************************
         *
         *       テーブル別のレシピテーブル一覧
         *
         ***************************************************/
        class SelectTbTable extends JTable {

            private SelectTbMdl model = null;
	        private boolean life            = false;

            // ---------- コンストラクタ -------------------
            //
            SelectTbTable(int[] tbl_no, String[] t_title){
                super();

                try{
                    setName("SelectTbTable");
                    setBounds(0, 0, 200, 200);
                    setAutoCreateColumnsFromModel(true);
                    setSelectionMode(ListSelectionModel.SINGLE_SELECTION);

                    setLocale(new Locale("ja","JP"));
                    setFont(new java.awt.Font("dialog", 0, 12));
                    setRowHeight(17);

                    model = new SelectTbMdl(tbl_no, t_title);
                    setModel(model);

                    DefaultTableColumnModel cmdl = (DefaultTableColumnModel)getColumnModel();
                    TableColumn  colum = null;
                    // テーブルNo
                    colum = cmdl.getColumn(0);
                    colum.setMaxWidth(40);
                    colum.setMinWidth(40);
                    colum.setWidth(40);
                    // 項目名
                    colum = cmdl.getColumn(1);
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
//@@                CZSystem.log("SelectTbTable","GroupeTable valueChanged [" +
//@@                            getSelectedRow() + "][" + getSelectedColumn() + "]");
            }


            //
            //
            //
            public void setData(int gr,int tbl){
//
//@@                CZSystem.log("CZControlTableCP","SelectTbTable setData() [" + gr + "][" + tbl + "]");
//                model.setValueAt(tbl_no, t_title);
//
//              setRowSelectionInterval(0,0);
//
//                Rectangle cellRect = getCellRect(0,0,false);
//                if(cellRect != null){
//                    scrollRectToVisible(cellRect);
//               }
//                repaint();
            }
        }

        /***************************************************
         *
         *       制御テーブルクラス：モデル
         *
         ***************************************************/
        public class SelectTbMdl extends AbstractTableModel {

            final   int TBL_COL     = 2;
            private int TBL_ROW     = 0;

            final String[] names = {" # " ,"項目"};

            private Object  data[][];

            SelectTbMdl(int[] tbl_no, String[] t_title){
                super();

				TBL_ROW = tbl_no.length;
				
				data = new Object[TBL_ROW][TBL_COL];
//                CZSystem.log("CZControlTableCP","setValueAt TBL_ROW[" + TBL_ROW  + "]");
//
//                for(int i = 0 ; i < TBL_ROW ; i++){
//                    CZSystem.log("CZControlTableCP","#[" + i + "] tbl_no[" + tbl_no[i]  + "][" + t_title[i] + "]");
//                }

                for(int i = 0 ; i < TBL_ROW ; i++){
                    data[i][0] = new Integer(tbl_no[i]);
                    data[i][1] = new String(t_title[i]);
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

///2011.04.12 Y.K END


        /***************************************************
         *
         *       グループ別のレシピテーブル一覧
         *
         ***************************************************/
        class GroupeTable extends JTable {

            private GrTblMdl model = null;

            // ---------- コンストラクタ -------------------
            //
            GroupeTable(){
                super();

                try{
                    setName("GroupeTable");
                    setBounds(0, 0, 200, 200);
                    setAutoCreateColumnsFromModel(true);
                    setSelectionMode(ListSelectionModel.SINGLE_SELECTION);

                    setLocale(new Locale("ja","JP"));
                    setFont(new java.awt.Font("dialog", 0, 12));
                    setRowHeight(17);

                    model = new GrTblMdl();
                    setModel(model);

                    DefaultTableColumnModel cmdl = (DefaultTableColumnModel)getColumnModel();
                    TableColumn  colum = null;
                    // レシピーNo
                    colum = cmdl.getColumn(0);
                    colum.setMaxWidth(40);
                    colum.setMinWidth(40);
                    colum.setWidth(40);
                    // タイトル
                    colum = cmdl.getColumn(1);
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
//@@                CZSystem.log("CZControlTableCP","GroupeTable valueChanged [" +
//@@                            getSelectedRow() + "][" + getSelectedColumn() + "]");
            }


            //
            //
            //
            public void setData(int gr,int tbl){

//@@                CZSystem.log("CZControlTableCP","GroupeTable setData() [" + gr + "][" + tbl + "]");
                CZSystemCtTitle title   = null;
                String          empty   = new String("");
                for(int i = 0 ; i < 999 ; i++){
                    g_table.setValueAt(empty,i,1);
                }

                for(int i = 0 ; i < table_title.size() ; i++){
                    title = (CZSystemCtTitle)table_title.elementAt(i);
                    if(gr == title.g_no){
                        g_table.setValueAt(title.title,title.r_no-1,1);
                    }
                }

                setRowSelectionInterval(tbl-1,tbl-1);

                Rectangle cellRect = getCellRect(tbl-1,0,false);
                if(cellRect != null){
                    scrollRectToVisible(cellRect);
                }
                repaint();
            }
        }

        /***************************************************
         *
         *       制御テーブルクラス：モデル
         *
         ***************************************************/
        public class GrTblMdl extends AbstractTableModel {

            final   int TBL_COL     = 2;
            private int TBL_ROW     = 999;

            final String[] names = {" # " ,"タイトル"};

            private Object  data[][];

            GrTblMdl(){
                super();

                data = new Object[TBL_ROW][TBL_COL];

                for(int i = 0 ; i < TBL_ROW ; i++){
                    data[i][0] = new Integer(i+1);
                    data[i][1] = new String("");
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

        /***************************************************
         *
         *   実行ボタン
         *
         ***************************************************/
        class SendButton implements ActionListener {
            public void actionPerformed(ActionEvent ev){

                if(!setSendStatus()){
                    Object msg[] = {"テーブルコピー",
                                "テーブル、設定者を",
                                "見直してください"};
                    errorMsg(msg);
                    return;
                }

                //Send
                String sendOp = op_name.getText();
                int idx = dst_ro_name.getSelectedIndex();
                String dst_ro = CZSystem.getRoName(idx);

///////2011.04.12 Y.K start

	            JViewport v;
	            SelectTbTable t;

	            v = panelFromTbl.getViewport();
	            t = (SelectTbTable)v.getView();
	            if(null == t)
				{
//					CZSystem.log("CZControlTableCP","get SelectTbTable Ng");
					return;
				}
	            int iRow_max = t.getRowCount();
//				CZSystem.log("CZControlTableCP","get row_max [" + iRow_max + "]");

				int intSelect_t;
				Integer Select_t;
				for (int iLp = 0; iLp < iRow_max; iLp++)
				{
	//.intValue()
					Select_t = (Integer)t.getValueAt(iLp, 0);
					intSelect_t = Select_t.intValue();
//					CZSystem.log("CZControlTableCP","get ValueAt [" + iLp + "][" + intSelect_t + "]");

	                if(!CZSystem.CZControlCopyTable(sendOp,dst_ro,groupe_no,recipie_no,send_recipie_no,intSelect_t)){

	                    dst_ro_name.setDefault();

	                    Object msg[] = {"テーブルコピー",
	                                    "コピーが失敗しました(" + intSelect_t + ")",
	                                    "管理者に問い合わせてください"};
	                    errorMsg(msg);
	                    return;
	                }
				}
///////2011.04.12 Y.K end

                dst_ro_name.setDefault();
                return ;
            }
        }

        /***************************************************
         *
         *   終了ボタン
         *
         ***************************************************/
        class CancelButton implements ActionListener {
            public void actionPerformed(ActionEvent ev){
                setVisible(false);
                dst_ro_name.releaseHaita();
            }
        }
    } /* public class TableCopy extends JDialog */

    /*******************************************************
     *
     *       設定者を入力するTextField
     *
     *******************************************************/
    public class TText extends JTextField {

        // ---------- コンストラクタ -----------------------
        //
        TText(){
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
    }   //public class TableCopy extends JDialog

    //////////////////////////////////////////////////////////////////////////////////////////////
    //
    //  大項目コピー用画面
    //
    //////////////////////////////////////////////////////////////////////////////////////////////
    public class T6LagCopy extends JDialog {

        private int destRcpNo = 0;

        private int grpNo     = 0;
        private int rcpNo     = 0;
        private int lagNo     = 0;

        private String  grpName = null;
        private String  rcpName = null;
        private String  lagName = null;

        private JLabel  srcRoName  = null;
        private JLabel  srcGrpName = null;
        private JLabel  srcRcpNo   = null;
        private JLabel  srcRcpName = null;
        private JLabel  srcLagNo   = null;
        private JLabel  srcLagName = null;

        private RoNo    dstRoName     = null;
        private RcpTable rcpTable     = null;

        private JButton cp_button     = null;

        private Vector  vRcpName      = null;

        private TText   op_name       = null;

        private int old_idx           = -1;

        //
        //
        //
        T6LagCopy(){
            setTitle("制御テーブル：大項目コピー");
            setSize(710,480);
            setResizable(false);
            setModal(true);
            getContentPane().setLayout(null);
            // 他基地参照機能    @20131021
            if(CZSystemDefine.REFERENCE_RUN == CZSystem.getRunLevel()){
                getContentPane().setBackground(CZSystemDefine.DEFAULT_REFERENCE_BACKGROUND_COL);
            }else{
                getContentPane().setBackground(CZSystemDefine.DEFAULT_BACKGROUND_COL);
            }

            addWindowListener(new WindowAdapter(){
                public void windowClosing(WindowEvent e){
                    winClose(e);
                }
            });

            JLabel  label   = null;
            label = new JLabel("コピー元",JLabel.CENTER);
            label.setBounds(20, 20, 100, 24);
            label.setLocale(new Locale("ja","JP"));
            label.setFont(new java.awt.Font("dialog", 0, 16));
            label.setBorder(new Flush3DBorder());
            label.setForeground(java.awt.Color.black);
            getContentPane().add(label);

            srcRoName = new JLabel(" ",JLabel.CENTER);
            srcRoName.setBounds(20, 54, 100, 24);
            srcRoName.setLocale(new Locale("ja","JP"));
            srcRoName.setFont(new java.awt.Font("dialog", 0, 16));
            srcRoName.setBorder(new Flush3DBorder());
            srcRoName.setForeground(java.awt.Color.black);
            getContentPane().add(srcRoName);

            srcGrpName = new JLabel(" ",JLabel.CENTER);
            srcGrpName.setBounds(140, 54, 100, 24);
            srcGrpName.setLocale(new Locale("ja","JP"));
            srcGrpName.setFont(new java.awt.Font("dialog", 0, 16));
            srcGrpName.setBorder(new Flush3DBorder());
            srcGrpName.setForeground(java.awt.Color.black);
            getContentPane().add(srcGrpName);

            srcRcpNo = new JLabel(" ",JLabel.CENTER);
            srcRcpNo.setBounds(20, 88, 100, 24);
            srcRcpNo.setLocale(new Locale("ja","JP"));
            srcRcpNo.setFont(new java.awt.Font("dialog", 0, 16));
            srcRcpNo.setBorder(new Flush3DBorder());
            srcRcpNo.setForeground(java.awt.Color.black);
            getContentPane().add(srcRcpNo);

            srcRcpName = new JLabel(" ",JLabel.CENTER);
            srcRcpName.setBounds(140, 88, 540, 24);
            srcRcpName.setLocale(new Locale("ja","JP"));
            srcRcpName.setFont(new java.awt.Font("dialog", 0, 16));
            srcRcpName.setBorder(new Flush3DBorder());
            srcRcpName.setForeground(java.awt.Color.black);
            getContentPane().add(srcRcpName);

            srcLagNo = new JLabel(" ",JLabel.CENTER);
            srcLagNo.setBounds(20, 112, 100, 24);
            srcLagNo.setLocale(new Locale("ja","JP"));
            srcLagNo.setFont(new java.awt.Font("dialog", 0, 16));
            srcLagNo.setBorder(new Flush3DBorder());
            srcLagNo.setForeground(java.awt.Color.black);
            getContentPane().add(srcLagNo);

            srcLagName = new JLabel(" ",JLabel.CENTER);
            srcLagName.setBounds(140, 112, 540, 24);
            srcLagName.setLocale(new Locale("ja","JP"));
            srcLagName.setFont(new java.awt.Font("dialog", 0, 16));
            srcLagName.setBorder(new Flush3DBorder());
            srcLagName.setForeground(java.awt.Color.black);
            getContentPane().add(srcLagName);

            label = new JLabel("コピー先",JLabel.CENTER);
            label.setBounds(20, 150, 100, 24);
            label.setLocale(new Locale("ja","JP"));
            label.setFont(new java.awt.Font("dialog", 0, 16));
            label.setBorder(new Flush3DBorder());
            label.setForeground(java.awt.Color.black);
            getContentPane().add(label);

            dstRoName = new RoNo();
            dstRoName.setBounds(20, 184, 100, 24);
            getContentPane().add(dstRoName);

            JScrollPane panel = null;
            rcpTable = new RcpTable();
            panel = new JScrollPane(rcpTable);
            panel.setBounds(140, 184, 540, 200);
            getContentPane().add(panel);

            label = new JLabel("設定者",JLabel.CENTER);
            label.setBounds(20, 410, 100, 24);
            label.setLocale(new Locale("ja","JP"));
            label.setFont(new java.awt.Font("dialog", 0, 16));
            label.setBorder(new Flush3DBorder());
            label.setForeground(java.awt.Color.black);
            getContentPane().add(label);

            // オペレータ名
            op_name = new TText();
            op_name.setBounds(120, 410, 140, 24);
            getContentPane().add(op_name);

            cp_button = new JButton();
            cp_button = new JButton("実  行");
            cp_button.setBounds(260, 410, 100, 24);
            cp_button.setLocale(new Locale("ja","JP"));
            cp_button.setFont(new java.awt.Font("dialog", 0, 18));
            cp_button.setBorder(new Flush3DBorder());
            cp_button.setForeground(java.awt.Color.black);
            cp_button.addActionListener(new SendButton());
            getContentPane().add(cp_button);

            // 終了
            JButton button = new JButton("終  了");
            button.setBounds(580, 410, 100, 24);
            button.setLocale(new Locale("ja","JP"));
            button.setFont(new java.awt.Font("dialog", 0, 18));
            button.setBorder(new Flush3DBorder());
            button.setForeground(java.awt.Color.black);
            button.addActionListener(new CancelButton());
            getContentPane().add(button);
        }

        //
        //
        //
        private void winClose(WindowEvent e){
            CZSystem.log("CZControlTableCP","T6LagCopy winClose() " + e);
            dstRoName.releaseHaita();
        }


        //
        //
        //
        public boolean setDefault(
            int gNo,String gName,
            int rNo,String rName,
            int lNo,String lName){

//@@            CZSystem.log("CZControlTableCP","T6LagCopy " +
//@@                               "gNo[" + gNo + "] gName[" + gName + "] " +
//@@                               "rNo[" + rNo + "] rName[" + rName +"]" +
//@@                               "lNo[" + lNo + "] lName[" + lName +"]");

            op_name.setText("");
            grpNo   = gNo;
            rcpNo   = rNo;
            grpName = gName;
            rcpName = rName;
            lagNo   = lNo;
            lagName = lName;

			String s = CZSystem.RoKetaChg(CZSystem.getRoName());	// 20050725 炉：表示桁数変更
            srcRoName.setText(s);
//            srcRoName.setText(CZSystem.getRoName());
            srcGrpName.setText(grpName);
            srcRcpNo.setText(new String(" " + rcpNo + " "));
            srcRcpName.setText(rcpName);
            srcLagNo.setText(new String(" " + lagNo + " "));
            srcLagName.setText(lagName);
            dstRoName.setDefault();
            return true;
        }


        //
        //
        //
        public boolean setSendStatus(){
            int idx = 0;
            idx = dstRoName.getSelectedIndex();
            if(0 > idx) return false;
            String sendOp = op_name.getText();
            if(1 > sendOp.length()) return false;

            destRcpNo = rcpTable.getSelectedRow() + 1;
            if(1 > destRcpNo) return false;

            String ro     = CZSystem.getRoName();
            String dstRo  = CZSystem.getRoName(idx);
            if( ro.equals(dstRo) && (rcpNo == destRcpNo)) return false;

//@@            CZSystem.log("CZControlTableCP","T6LagCopy grp_no[" + grpNo + "]");
//@@            CZSystem.log("CZControlTableCP","T6LagCopy ro    [" + ro + "]->[" + dstRo + "]");
//@@            CZSystem.log("CZControlTableCP","T6LagCopy rec_no[" + rcpNo + "]->[" + destRcpNo + "]");

            return true;
        }


        //
        // メッセージの表示
        //
        private boolean errorMsg(Object msg[]){
            JOptionPane.showMessageDialog(null,msg,
                    "制御テーブル：大項目コピー",
                    JOptionPane.ERROR_MESSAGE);
            return true;
        }


        //
        //
        //
        public class RoNo extends JComboBox {

            RoNo(){
                super();
                try{
                    setName("JComboBox1");
                    setFont(new java.awt.Font("dialog", 0, 16));
                    Vector ro = CZSystem.getRoNameList();
                    if(null == ro){
                        CZSystem.exit(0,"Not Ro No");
                    }
                    for(int i = 0 ; ro.size() > i ; i++){
						String s = CZSystem.RoKetaChg((String)ro.elementAt(i));	// 20050725 炉：表示桁数変更
                        addItem(s);
//                        addItem((String)ro.elementAt(i));
                    }
                    setForeground(java.awt.Color.black);
                    setBackground(java.awt.Color.lightGray);
					setFocusable(false);	/* 2007.08.22 */
                    addActionListener(new ChgRoNo());
//@@                    CZSystem.log("CZControlTableCP","T6LagCopy new RoNo()");
                }
                catch (Throwable e) {
                    CZSystem.handleException(e);
                }
            }


            //
            //
            //
            public void setDefault(){

                int idx = getSelectedIndex();
                old_idx = idx;      //@@@

//@@                CZSystem.log("CZControlTableCP","T6LagCopy setDefault() RoNo[" +
//@@                                idx + "][" + CZSystem.getRoName(idx) + "]" );
				String s = CZSystem.RoKetaChg(CZSystem.getRoName(idx));	// 20050725 炉：表示桁数変更

                vRcpName = null;
                vRcpName  = CZSystem.getCtTitle(idx);
                if(null != vRcpName){
                    rcpTable.setData(grpNo,rcpNo);
                }
                if(getHaita(idx)){
                    cp_button.setEnabled(true);
                }
                else {
                    cp_button.setEnabled(false);
                    Object msg[] = {"大項目コピー",
                        new String(s + "は"),
//                        new String(CZSystem.getRoName(idx) + "は"),
                        "制御盤、他の端末で修正中です"};
                    errorMsg(msg);
                }
            }


            //
            // 画面消去時の排他開放
            //
            public void releaseHaita(){
                int idx = getSelectedIndex();
//@@                CZSystem.log("CZControlTableCP","T6LagCopy releaseHaita() 排他[" + idx + "]開放");
                putHaita(idx);
            }


            //
            //
            //
            class ChgRoNo implements ActionListener {
                public void actionPerformed(ActionEvent e){

                    RoNo obj = (RoNo)e.getSource();
                    if(-1 == old_idx){
//@@                        CZSystem.log("CZControlTableCP","T6LagCopy ChgRoNo 排他１回目");
                    }
                    else {
                        putHaita(old_idx);
//@@                        CZSystem.log("CZControlTableCP","T6LagCopy ChgRoNo 排他[" +
//@@                                old_idx + "]->[" + obj.getSelectedIndex() + "]");
                    }
                    obj.setDefault();
                    old_idx = obj.getSelectedIndex();
                }
            }

        }

        /***************************************************
         *
         *       グループ別のレシピテーブル一覧
         *
         ***************************************************/
        class RcpTable extends JTable {

            private RcpTblMdl model = null;

            RcpTable(){
                super();
                try{
                    setName("RcpTable");
                    setBounds(0, 0, 200, 200);
                    setAutoCreateColumnsFromModel(true);
                    setSelectionMode(ListSelectionModel.SINGLE_SELECTION);
                    setLocale(new Locale("ja","JP"));
                    setFont(new java.awt.Font("dialog", 0, 12));
                    setRowHeight(17);
                    model = new RcpTblMdl();
                    setModel(model);
                    DefaultTableColumnModel cmdl = (DefaultTableColumnModel)getColumnModel();
                    TableColumn  colum = null;
                    // レシピーNo
                    colum = cmdl.getColumn(0);
                    colum.setMaxWidth(40);
                    colum.setMinWidth(40);
                    colum.setWidth(40);
                    // タイトル
                    colum = cmdl.getColumn(1);
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
//@@                CZSystem.log("CZControlTableCp","RcpTable valueChanged [" +
//@@                            getSelectedRow() + "][" + getSelectedColumn() + "]");
            }

            //
            //
            //
            public void setData(int gr,int tbl){

//@@                CZSystem.log("CZControlTableCp","RcpTable setData() [" + gr + "][" + tbl + "]");
                CZSystemCtTitle rcpName   = null;
                String          empty   = new String("");

                for(int i = 0 ; i < 999 ; i++){
                    rcpTable.setValueAt(empty,i,1);
                }

                for(int i = 0 ; i < vRcpName.size() ; i++){
                    rcpName = (CZSystemCtTitle)vRcpName.elementAt(i);
                    if(gr == rcpName.g_no){
                        rcpTable.setValueAt(rcpName.title,rcpName.r_no-1,1);
                    }
                }
                setRowSelectionInterval(tbl-1,tbl-1);
                Rectangle cellRect = getCellRect(tbl-1,0,false);
                if(cellRect != null){
                    scrollRectToVisible(cellRect);
                }
                repaint();
            }
        }

        /***************************************************
         *
         *       制御テーブルクラス：モデル
         *
         ***************************************************/
        public class RcpTblMdl extends AbstractTableModel {

            final   int TBL_COL     = 2;
            private int TBL_ROW     = 999;
            final String[] names = {" # " ,"タイトル"};
            private Object  data[][];

            //
            RcpTblMdl(){
                super();
                data = new Object[TBL_ROW][TBL_COL];
                for(int i = 0 ; i < TBL_ROW ; i++){
                    data[i][0] = new Integer(i+1);
                    data[i][1] = new String("");
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

        /***************************************************
         *
         *   実行ボタン
         *
         ***************************************************/
        class SendButton implements ActionListener {
            public void actionPerformed(ActionEvent ev){

                if(!setSendStatus()){
                    Object msg[] = {"大項目コピー",
                                    "大項目、設定者を",
                                    "見直してください"};
                    errorMsg(msg);
                    return;
                 }
                //Send
                String sendOp = op_name.getText();
                int idx = dstRoName.getSelectedIndex();
                String dstRo = CZSystem.getRoName(idx);
                if(!CZSystem.CZControlCopyLagName(sendOp,dstRo,grpNo,rcpNo,destRcpNo,lagNo)){
                    dstRoName.setDefault();
                    Object msg[] = {"大項目コピー",
                                    "コピーが失敗しました",
                                    "管理者に問い合わせてください"};
                    errorMsg(msg);
                    return;
                }
                dstRoName.setDefault();
                return ;
            }
        }



        /***************************************************
         *
         *   終了ボタン
         *
         ***************************************************/
        class CancelButton implements ActionListener {
            public void actionPerformed(ActionEvent ev){
                setVisible(false);
                dstRoName.releaseHaita();
            }
        }
    } /* public class T6LagCopy extends JDialog */


    //////////////////////////////////////////////////////////////////////////////////////////////
    //
    //  中項目コピー用画面
    //
    //////////////////////////////////////////////////////////////////////////////////////////////
    public class T6MidCopy extends JDialog {

        private int destLagNo       = 0;

        private int grpNo           = 0;
        private int rcpNo           = 0;
        private int lagNo           = 0;
        private int midNo           = 0;
        private String  grpName     = null;
        private String  rcpName     = null;
        private String  lagName     = null;
        private String  midName     = null;

        private JLabel  srcRoName   = null;
        private JLabel  srcGrpNo    = null;
        private JLabel  srcGrpName  = null;
        private JLabel  srcRcpNo    = null;
        private JLabel  srcRcpName  = null;

        private JLabel  srcLagNo    = null;
        private JLabel  srcLagName  = null;
        private JLabel  srcMidNo    = null;
        private JLabel  srcMidName  = null;

        private RoNo    dstRoName       = null;
        private T6LagTable t6LagTable   = null;

        private JButton cp_button       = null;
//2011.04.14 Y.K レシピタイトルに変更
//        private Vector  vLagName        = null;
        private Vector  vRcpName      = null;

        private TText   op_name         = null;

        private int old_idx             = -1;

        //
        //
        //
        T6MidCopy(){
            setTitle("制御テーブル：中項目コピー");
            setSize(710,480);
            setResizable(false);
            setModal(true);
            getContentPane().setLayout(null);
            // 他基地参照機能    @20131021
            if(CZSystemDefine.REFERENCE_RUN == CZSystem.getRunLevel()){
                getContentPane().setBackground(CZSystemDefine.DEFAULT_REFERENCE_BACKGROUND_COL);
            }else{
                getContentPane().setBackground(CZSystemDefine.DEFAULT_BACKGROUND_COL);
            }

            addWindowListener(new WindowAdapter(){
                public void windowClosing(WindowEvent e){
                    winClose(e);
                }
            });

            JLabel  label   = null;
            label = new JLabel("コピー元",JLabel.CENTER);
            label.setBounds(20, 20, 100, 24);
            label.setLocale(new Locale("ja","JP"));
            label.setFont(new java.awt.Font("dialog", 0, 16));
            label.setBorder(new Flush3DBorder());
            label.setForeground(java.awt.Color.black);
            getContentPane().add(label);

            srcRoName = new JLabel(" ",JLabel.CENTER);
            srcRoName.setBounds(20, 54, 100, 24);
            srcRoName.setLocale(new Locale("ja","JP"));
            srcRoName.setFont(new java.awt.Font("dialog", 0, 16));
            srcRoName.setBorder(new Flush3DBorder());
            srcRoName.setForeground(java.awt.Color.black);
            getContentPane().add(srcRoName);

            srcGrpName = new JLabel(" ",JLabel.CENTER);
            srcGrpName.setBounds(140, 54, 100, 24);
            srcGrpName.setLocale(new Locale("ja","JP"));
            srcGrpName.setFont(new java.awt.Font("dialog", 0, 16));
            srcGrpName.setBorder(new Flush3DBorder());
            srcGrpName.setForeground(java.awt.Color.black);
            getContentPane().add(srcGrpName);

            srcRcpNo = new JLabel(" ",JLabel.CENTER);
            srcRcpNo.setBounds(20, 88, 100, 24);
            srcRcpNo.setLocale(new Locale("ja","JP"));
            srcRcpNo.setFont(new java.awt.Font("dialog", 0, 16));
            srcRcpNo.setBorder(new Flush3DBorder());
            srcRcpNo.setForeground(java.awt.Color.black);
            getContentPane().add(srcRcpNo);

            srcRcpName = new JLabel(" ",JLabel.CENTER);
            srcRcpName.setBounds(140, 88, 540, 24);
            srcRcpName.setLocale(new Locale("ja","JP"));
            srcRcpName.setFont(new java.awt.Font("dialog", 0, 16));
            srcRcpName.setBorder(new Flush3DBorder());
            srcRcpName.setForeground(java.awt.Color.black);
            getContentPane().add(srcRcpName);

            srcLagNo = new JLabel(" ",JLabel.CENTER);
            srcLagNo.setBounds(20, 112, 100, 24);
            srcLagNo.setLocale(new Locale("ja","JP"));
            srcLagNo.setFont(new java.awt.Font("dialog", 0, 16));
            srcLagNo.setBorder(new Flush3DBorder());
            srcLagNo.setForeground(java.awt.Color.black);
            getContentPane().add(srcLagNo);

            srcLagName = new JLabel(" ",JLabel.CENTER);
            srcLagName.setBounds(140, 112, 540, 24);
            srcLagName.setLocale(new Locale("ja","JP"));
            srcLagName.setFont(new java.awt.Font("dialog", 0, 16));
            srcLagName.setBorder(new Flush3DBorder());
            srcLagName.setForeground(java.awt.Color.black);
            getContentPane().add(srcLagName);

            srcMidNo = new JLabel(" ",JLabel.CENTER);
            srcMidNo.setBounds(20, 136, 100, 24);
            srcMidNo.setLocale(new Locale("ja","JP"));
            srcMidNo.setFont(new java.awt.Font("dialog", 0, 16));
            srcMidNo.setBorder(new Flush3DBorder());
            srcMidNo.setForeground(java.awt.Color.black);
            getContentPane().add(srcMidNo);

            srcMidName = new JLabel(" ",JLabel.CENTER);
            srcMidName.setBounds(140, 136, 540, 24);
            srcMidName.setLocale(new Locale("ja","JP"));
            srcMidName.setFont(new java.awt.Font("dialog", 0, 16));
            srcMidName.setBorder(new Flush3DBorder());
            srcMidName.setForeground(java.awt.Color.black);
            getContentPane().add(srcMidName);

            label = new JLabel("コピー先",JLabel.CENTER);
            label.setBounds(20, 184, 100, 24);
            label.setLocale(new Locale("ja","JP"));
            label.setFont(new java.awt.Font("dialog", 0, 16));
            label.setBorder(new Flush3DBorder());
            label.setForeground(java.awt.Color.black);
            getContentPane().add(label);

            dstRoName = new RoNo();
            dstRoName.setBounds(20, 214, 100, 24);
            getContentPane().add(dstRoName);

            JScrollPane panel = null;
            t6LagTable = new T6LagTable();
            panel = new JScrollPane(t6LagTable);
            panel.setBounds(140, 184, 540, 200);
            getContentPane().add(panel);

            label = new JLabel("設定者",JLabel.CENTER);
            label.setBounds(20, 410, 100, 24);
            label.setLocale(new Locale("ja","JP"));
            label.setFont(new java.awt.Font("dialog", 0, 16));
            label.setBorder(new Flush3DBorder());
            label.setForeground(java.awt.Color.black);
            getContentPane().add(label);

            // オペレータ名
            op_name = new TText();
            op_name.setBounds(120, 410, 140, 24);
            getContentPane().add(op_name);

            cp_button = new JButton();
            cp_button = new JButton("実  行");
            cp_button.setBounds(260, 410, 100, 24);
            cp_button.setLocale(new Locale("ja","JP"));
            cp_button.setFont(new java.awt.Font("dialog", 0, 18));
            cp_button.setBorder(new Flush3DBorder());
            cp_button.setForeground(java.awt.Color.black);
            cp_button.addActionListener(new SendButton());
            getContentPane().add(cp_button);

            // 終了
            JButton button = new JButton("終  了");
            button.setBounds(580, 410, 100, 24);
            button.setLocale(new Locale("ja","JP"));
            button.setFont(new java.awt.Font("dialog", 0, 18));
            button.setBorder(new Flush3DBorder());
            button.setForeground(java.awt.Color.black);
            button.addActionListener(new CancelButton());
            getContentPane().add(button);
        }

        //
        //
        //
        private void winClose(WindowEvent e){
            CZSystem.log("CZControlTableCp","T6MidTable winClose() " + e);
            dstRoName.releaseHaita();
        }


        //
        //
        //
        public boolean setDefault(
                        int g_no, String g_name,
                        int r_no, String r_name,
                        int l_no, String l_name,
                        int m_no, String m_name ){

//@@            CZSystem.log("CZControlTableCp","T6MidTable "+
//@@                                "grp_no[" + g_no + "] g_name[" + g_name + "]" +
//@@                                "rec_no[" + r_no + "] r_name[" + r_name + "]" +
//@@                                "lag_no[" + l_no + "] l_name[" + l_name + "]" +
//@@                                "mid_no[" + m_no + "] m_name[" + m_name + "]" );

            op_name.setText("");
            grpNo    = g_no;
            rcpNo    = r_no;
            lagNo    = l_no;
            midNo    = m_no;

            grpName  = g_name;
            rcpName  = r_name;
            lagName  = l_name;
            midName  = m_name;

			String s = CZSystem.RoKetaChg(CZSystem.getRoName());	// 20050725 炉：表示桁数変更
            srcRoName.setText(s);
//            srcRoName.setText(CZSystem.getRoName());
            srcGrpName.setText(grpName);
            srcRcpNo.setText(new String(" " + rcpNo + " "));
            srcRcpName.setText(rcpName);
            srcLagNo.setText(new String(" " + lagNo + " "));
            srcLagName.setText(lagName);
            srcMidNo.setText(new String(" " + midNo + " "));
            srcMidName.setText(midName);
            dstRoName.setDefault();
            return true;
        }


        //
        //
        //
        public boolean setSendStatus(){
            int idx = 0;
            idx = dstRoName.getSelectedIndex();
            if(0 > idx) return false;
            String sendOp = op_name.getText();
            if(1 > sendOp.length()) return false;

            destLagNo = t6LagTable.getSelectedRow() + 1;
            if(1 > destLagNo) return false;

            String ro     = CZSystem.getRoName();
            String dstRo  = CZSystem.getRoName(idx);
//2003.11.12 syusei
//            if( ro.equals(dstRo) && (lagNo == destLagNo)) return false;
            if( ro.equals(dstRo) && (rcpNo == destLagNo)) return false;
//2003.11.12 syusei

//@@            CZSystem.log("CZControlTableCp","T6MidTable grp_no[" + grpNo  + "]");
//@@            CZSystem.log("CZControlTableCp","T6MidTable ro    [" + ro + "]->[" + dstRo + "]");
//@@            CZSystem.log("CZControlTableCp","T6MidTable Lag_no[" + lagNo + "]->[" + destLagNo + "]");

            return true;
        }

        //
        // メッセージの表示
        //
        private boolean errorMsg(Object msg[]){
            JOptionPane.showMessageDialog(null,msg,
                    "制御テーブル：中項目コピー",
                    JOptionPane.ERROR_MESSAGE);
            return true;
        }

        //
        //
        //
        public class RoNo extends JComboBox {

            RoNo(){
                super();
                try{
                    setName("JComboBox1");
                    setFont(new java.awt.Font("dialog", 0, 16));
                    Vector ro = CZSystem.getRoNameList();
                    if(null == ro){
                        CZSystem.exit(0,"Not Ro No");
                    }
                    for(int i = 0 ; ro.size() > i ; i++){
						String s = CZSystem.RoKetaChg((String)ro.elementAt(i));	// 20050725 炉：表示桁数変更
                        addItem(s);
//                        addItem((String)ro.elementAt(i));
                    }
                    setForeground(java.awt.Color.black);
                    setBackground(java.awt.Color.lightGray);
					setFocusable(false);	/* 2007.08.22 */
                    addActionListener(new ChgRoNo());
//@@                    CZSystem.log("CZControlTableCp","T6MidTable new RoNo()");
                }
                catch (Throwable e) {
                    CZSystem.handleException(e);
                }
            }

            //
            //
            //
            public void setDefault(){

                int idx = getSelectedIndex();
                old_idx = idx;      //@@@
//@@                CZSystem.log("CZControlTableCP","T6LagCopy setDefault() RoNo[" +
//@@                                    idx + "][" + CZSystem.getRoName(idx) + "]" );
				String s = CZSystem.RoKetaChg(CZSystem.getRoName(idx));	// 20050725 炉：表示桁数変更

//2011.04.14 Y.K レシピタイトルに修正
//                vLagName = null;
//                vLagName    = CZSystem.getCtT6Lag(grpNo,rcpNo);
//                if(null != vLagName){
                vRcpName = null;
                vRcpName  = CZSystem.getCtTitle(idx);
                if(null != vRcpName){
//@@                    System.out.println("t6LagTable.setData(grpNo="+ grpNo + ":rcpNo=" + rcpNo + ":idx="+ idx);
//2011.04.14 Y.K レシピタイトルに修正
//                    t6LagTable.setData(grpNo,rcpNo,idx+1);
                    t6LagTable.setData(grpNo,rcpNo);
                }

                if(getHaita(idx)){
                    cp_button.setEnabled(true);
                }
                else {
                    cp_button.setEnabled(false);
                    Object msg[] = {"中項目コピー",
                        new String(s + "は"),
//                        new String(CZSystem.getRoName(idx) + "は"),
                        "制御盤、他の端末で修正中です"};
                    errorMsg(msg);
                }
            }

            //
            // 画面消去時の排他開放
            //
            public void releaseHaita(){
                int idx = getSelectedIndex();
//@@                CZSystem.log("CZControlTableCp","T6MidTable releaseHaita() 排他[" + idx + "]開放");
                putHaita(idx);
            }

            //
            //
            //
            class ChgRoNo implements ActionListener {
                public void actionPerformed(ActionEvent e){

                    RoNo obj = (RoNo)e.getSource();
                    if(-1 == old_idx){
//@@                        CZSystem.log("CZControlTableCp","T6MidTable ChgRoNo()排他１回目");
                    }
                    else {
                        putHaita(old_idx);
//@@                        CZSystem.log("CZControlTableCp","T6MidTable ChgRoNo()排他[" +
//@@                                    old_idx + "]->[" + obj.getSelectedIndex() + "]");
                    }
                    obj.setDefault();
                    old_idx = obj.getSelectedIndex();
                }
            }

        }


        /***************************************************
         *
         *       グループ別のレシピ別大項目テーブル一覧
         *
         ***************************************************/
        class T6LagTable extends JTable {

            private T6LagTblMdl model = null;

            T6LagTable(){
                super();
                try{
                    setName("LargeTable");
                    setBounds(0, 0, 200, 200);
                    setAutoCreateColumnsFromModel(true);
                    setSelectionMode(ListSelectionModel.SINGLE_SELECTION);
                    setLocale(new Locale("ja","JP"));
                    setFont(new java.awt.Font("dialog", 0, 12));
                    setRowHeight(17);
                    model = new T6LagTblMdl();
                    setModel(model);
                    DefaultTableColumnModel cmdl = (DefaultTableColumnModel)getColumnModel();
                    TableColumn  colum = null;
                    // レシピーNo　2011.04.14 Y.K 大項目No=>レシピNoに変更
                    colum = cmdl.getColumn(0);
                    colum.setMaxWidth(40);
                    colum.setMinWidth(40);
                    colum.setWidth(40);
                    //  レシピーNo　2011.04.14 Y.K 大項目=>レシピNoに変更
                    colum = cmdl.getColumn(1);
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
//@@                CZSystem.log("CZControlTableCp","T6LagTable valueChanged [" +
//@@                                getSelectedRow() + "][" + getSelectedColumn() + "]");
            }

            //
            //
            //
//2011.04.14 Y.K Start
//            public void setData(int gr,int rcp,int lag){
            public void setData(int gr,int rcp){
//2011.04.14 Y.K End

//@@                CZSystem.log("CZControlTableCp","T6LagTable setData() [" +
//@@                                gr + "][" + rcp + "][" + lag + "]");
//2011.04.14 Y.K Start
//                CZSystemCtT6LagName t6Name   = null;
                CZSystemCtTitle rcpName   = null;
//2011.04.14 Y.K End
                String          empty   = new String("");
                for(int i = 0 ; i < 999 ; i++){
                    t6LagTable.setValueAt(empty,i,1);
                }

//2011.04.14 Y.K Start
                for(int i = 0 ; i < vRcpName.size() ; i++){
                    rcpName = (CZSystemCtTitle)vRcpName.elementAt(i);
                    if (gr == rcpName.g_no) {
                        t6LagTable.setValueAt(rcpName.title,rcpName.r_no-1,1);
                    }
                }
//                if ( 0 < rcp ) {
                    setRowSelectionInterval(rcp-1,rcp-1);
                    Rectangle cellRect = getCellRect(rcp-1,0,false);
                    if(cellRect != null){
                        scrollRectToVisible(cellRect);
                    }
//                }
//2011.04.14 Y.K End
                repaint();
            }
        }

        /***************************************************
         *
         *       制御テーブル(大項目)クラス：モデル
         *
         ***************************************************/
        public class T6LagTblMdl extends AbstractTableModel {

            final   int TBL_COL     = 2;
            private int TBL_ROW     = 999;
            final String[] names = {" # " ,"タイトル"};
            private Object  data[][];

            //
            T6LagTblMdl(){
                super();
                data = new Object[TBL_ROW][TBL_COL];
                for(int i = 0 ; i < TBL_ROW ; i++){
                    data[i][0] = new Integer(i+1);
                    data[i][1] = new String("");
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

        /***************************************************
         *
         *   実行ボタン
         *
         ***************************************************/
        class SendButton implements ActionListener {
            public void actionPerformed(ActionEvent ev){

                if(!setSendStatus()){
                    Object msg[] = {"中項目コピー",
                                    "中項目、設定者を",
                                    "見直してください"};
                    errorMsg(msg);
                    return;
                 }
                //Send
                String sendOp = op_name.getText();
                int idx = dstRoName.getSelectedIndex();
                String dstRo = CZSystem.getRoName(idx);
//2003.11.12 syusei
//                if(!CZSystem.CZControlCopyMidName(sendOp,dstRo,grpNo,rcpNo,rcpNo,lagNo,destLagNo,midNo)){
                if(!CZSystem.CZControlCopyMidName(sendOp,dstRo,grpNo,rcpNo,destLagNo,lagNo,lagNo,midNo)){
//2003.11.12 syusei
                    dstRoName.setDefault();
                    Object msg[] = {"中項目コピー",
                                    "コピーが失敗しました",
                                    "管理者に問い合わせてください"};
                    errorMsg(msg);
                    return;
                }
                dstRoName.setDefault();
                return ;
            }
        }

        /***************************************************
         *
         *   終了ボタン
         *
         ***************************************************/
        class CancelButton implements ActionListener {
            public void actionPerformed(ActionEvent ev){
                setVisible(false);
                dstRoName.releaseHaita();
            }
        }
    } /* public class T6MidCopy extends JDialog */

    //////////////////////////////////////////////////////////////////////////////////////////////
    //
    //  Ｔ６項目コピー用画面
    //
    public class T6ItemCopy extends JDialog {

        private int destMidNo       = 0;

        private int grpNo           = 0;
        private int rcpNo           = 0;
        private int lagNo           = 0;
        private int midNo           = 0;
        private int itmNo           = 0;
        private String  grpName     = null;
        private String  rcpName     = null;
        private String  lagName     = null;
        private String  midName     = null;
        private String  itmName     = null;

        private JLabel  srcRoName   = null;
        private JLabel  srcGrpNo    = null;
        private JLabel  srcGrpName  = null;
        private JLabel  srcRcpNo    = null;
        private JLabel  srcRcpName  = null;

        private JLabel  srcLagNo    = null;
        private JLabel  srcLagName  = null;
        private JLabel  srcMidNo    = null;
        private JLabel  srcMidName  = null;
        private JLabel  srcItmNo    = null;
        private JLabel  srcItmName  = null;

        private RoNo    dstRoName       = null;
        private T6MidTable t6MidTable   = null;

        private JButton cp_button       = null;
//2011.04.14 Y.K レシピタイトルに変更
//        private Vector  vMidName        = null;
        private Vector  vRcpName      = null;

        private TText   op_name         = null;

        private int old_idx             = -1;

        //
        //
        //
        T6ItemCopy(){
            setTitle("制御テーブル：Ｔ６項目コピー");
            setSize(710,520);
            setResizable(false);
            setModal(true);
            getContentPane().setLayout(null);
            // 他基地参照機能    @20131021
            if(CZSystemDefine.REFERENCE_RUN == CZSystem.getRunLevel()){
                getContentPane().setBackground(CZSystemDefine.DEFAULT_REFERENCE_BACKGROUND_COL);
            }else{
                getContentPane().setBackground(CZSystemDefine.DEFAULT_BACKGROUND_COL);
            }

            addWindowListener(new WindowAdapter(){
                public void windowClosing(WindowEvent e){
                    winClose(e);
                }
            });

            JLabel  label   = null;
            label = new JLabel("コピー元",JLabel.CENTER);
            label.setBounds(20, 20, 100, 24);
            label.setLocale(new Locale("ja","JP"));
            label.setFont(new java.awt.Font("dialog", 0, 16));
            label.setBorder(new Flush3DBorder());
            label.setForeground(java.awt.Color.black);
            getContentPane().add(label);

            srcRoName = new JLabel(" ",JLabel.CENTER);
            srcRoName.setBounds(20, 54, 100, 24);
            srcRoName.setLocale(new Locale("ja","JP"));
            srcRoName.setFont(new java.awt.Font("dialog", 0, 16));
            srcRoName.setBorder(new Flush3DBorder());
            srcRoName.setForeground(java.awt.Color.black);
            getContentPane().add(srcRoName);

            srcGrpName = new JLabel(" ",JLabel.CENTER);
            srcGrpName.setBounds(140, 54, 100, 24);
            srcGrpName.setLocale(new Locale("ja","JP"));
            srcGrpName.setFont(new java.awt.Font("dialog", 0, 16));
            srcGrpName.setBorder(new Flush3DBorder());
            srcGrpName.setForeground(java.awt.Color.black);
            getContentPane().add(srcGrpName);

            srcRcpNo = new JLabel(" ",JLabel.CENTER);
            srcRcpNo.setBounds(20, 88, 100, 24);
            srcRcpNo.setLocale(new Locale("ja","JP"));
            srcRcpNo.setFont(new java.awt.Font("dialog", 0, 16));
            srcRcpNo.setBorder(new Flush3DBorder());
            srcRcpNo.setForeground(java.awt.Color.black);
            getContentPane().add(srcRcpNo);

            srcRcpName = new JLabel(" ",JLabel.CENTER);
            srcRcpName.setBounds(140, 88, 540, 24);
            srcRcpName.setLocale(new Locale("ja","JP"));
            srcRcpName.setFont(new java.awt.Font("dialog", 0, 16));
            srcRcpName.setBorder(new Flush3DBorder());
            srcRcpName.setForeground(java.awt.Color.black);
            getContentPane().add(srcRcpName);

            srcLagNo = new JLabel(" ",JLabel.CENTER);
            srcLagNo.setBounds(20, 112, 100, 24);
            srcLagNo.setLocale(new Locale("ja","JP"));
            srcLagNo.setFont(new java.awt.Font("dialog", 0, 16));
            srcLagNo.setBorder(new Flush3DBorder());
            srcLagNo.setForeground(java.awt.Color.black);
            getContentPane().add(srcLagNo);

            srcLagName = new JLabel(" ",JLabel.CENTER);
            srcLagName.setBounds(140, 112, 540, 24);
            srcLagName.setLocale(new Locale("ja","JP"));
            srcLagName.setFont(new java.awt.Font("dialog", 0, 16));
            srcLagName.setBorder(new Flush3DBorder());
            srcLagName.setForeground(java.awt.Color.black);
            getContentPane().add(srcLagName);

            srcMidNo = new JLabel(" ",JLabel.CENTER);
            srcMidNo.setBounds(20, 136, 100, 24);
            srcMidNo.setLocale(new Locale("ja","JP"));
            srcMidNo.setFont(new java.awt.Font("dialog", 0, 16));
            srcMidNo.setBorder(new Flush3DBorder());
            srcMidNo.setForeground(java.awt.Color.black);
            getContentPane().add(srcMidNo);

            srcMidName = new JLabel(" ",JLabel.CENTER);
            srcMidName.setBounds(140, 136, 540, 24);
            srcMidName.setLocale(new Locale("ja","JP"));
            srcMidName.setFont(new java.awt.Font("dialog", 0, 16));
            srcMidName.setBorder(new Flush3DBorder());
            srcMidName.setForeground(java.awt.Color.black);
            getContentPane().add(srcMidName);

            srcItmNo = new JLabel(" ",JLabel.CENTER);
            srcItmNo.setBounds(20, 160, 100, 24);
            srcItmNo.setLocale(new Locale("ja","JP"));
            srcItmNo.setFont(new java.awt.Font("dialog", 0, 16));
            srcItmNo.setBorder(new Flush3DBorder());
            srcItmNo.setForeground(java.awt.Color.black);
            getContentPane().add(srcItmNo);

            srcItmName = new JLabel(" ",JLabel.CENTER);
            srcItmName.setBounds(140, 160, 540, 24);
            srcItmName.setLocale(new Locale("ja","JP"));
            srcItmName.setFont(new java.awt.Font("dialog", 0, 16));
            srcItmName.setBorder(new Flush3DBorder());
            srcItmName.setForeground(java.awt.Color.black);
            getContentPane().add(srcItmName);

            label = new JLabel("コピー先",JLabel.CENTER);
            label.setBounds(20, 214, 100, 24);
            label.setLocale(new Locale("ja","JP"));
            label.setFont(new java.awt.Font("dialog", 0, 16));
            label.setBorder(new Flush3DBorder());
            label.setForeground(java.awt.Color.black);
            getContentPane().add(label);

            dstRoName = new RoNo();
            dstRoName.setBounds(20, 244, 100, 24);
            getContentPane().add(dstRoName);

            JScrollPane panel = null;
            t6MidTable = new T6MidTable();
            panel = new JScrollPane(t6MidTable);
            panel.setBounds(140, 214, 540, 200);
            getContentPane().add(panel);

            label = new JLabel("設定者",JLabel.CENTER);
            label.setBounds(20, 440, 100, 24);
            label.setLocale(new Locale("ja","JP"));
            label.setFont(new java.awt.Font("dialog", 0, 16));
            label.setBorder(new Flush3DBorder());
            label.setForeground(java.awt.Color.black);
            getContentPane().add(label);

            // オペレータ名
            op_name = new TText();
            op_name.setBounds(120, 440, 140, 24);
            getContentPane().add(op_name);

            cp_button = new JButton();
            cp_button = new JButton("実  行");
            cp_button.setBounds(260, 440, 100, 24);
            cp_button.setLocale(new Locale("ja","JP"));
            cp_button.setFont(new java.awt.Font("dialog", 0, 18));
            cp_button.setBorder(new Flush3DBorder());
            cp_button.setForeground(java.awt.Color.black);
            cp_button.addActionListener(new SendButton());
            getContentPane().add(cp_button);

            // 終了
            JButton button = new JButton("終  了");
            button.setBounds(580, 440, 100, 24);
            button.setLocale(new Locale("ja","JP"));
            button.setFont(new java.awt.Font("dialog", 0, 18));
            button.setBorder(new Flush3DBorder());
            button.setForeground(java.awt.Color.black);
            button.addActionListener(new CancelButton());
            getContentPane().add(button);
        }

        //
        //
        //
        private void winClose(WindowEvent e){
            CZSystem.log("CZControlTableCp","T6ItemCopy winClose() " + e);
            dstRoName.releaseHaita();
        }

        //
        //
        //
        public boolean setDefault(
                        int g_no, String g_name,
                        int r_no, String r_name,
                        int l_no, String l_name,
                        int m_no, String m_name,
                        int i_no, String i_name ){

//@@            CZSystem.log("CZControlTableCp","T6ItemCopy "+
//@@                                "grp_no[" + g_no + "] g_name[" + g_name + "]" +
//@@                                "rec_no[" + r_no + "] r_name[" + r_name + "]" +
//@@                                "lag_no[" + l_no + "] l_name[" + l_name + "]" +
//@@                                "mid_no[" + m_no + "] m_name[" + m_name + "]" +
//@@                                "itm_no[" + i_no + "] i_name[" + i_name + "]" );

            op_name.setText("");
            grpNo    = g_no;
            rcpNo    = r_no;
            lagNo    = l_no;
            midNo    = m_no;
            itmNo    = i_no;

            grpName  = g_name;
            rcpName  = r_name;
            lagName  = l_name;
            midName  = m_name;
            itmName  = i_name;

			String s = CZSystem.RoKetaChg(CZSystem.getRoName());	// 20050725 炉：表示桁数変更
            srcRoName.setText(s);
//            srcRoName.setText(CZSystem.getRoName());
            srcGrpName.setText(grpName);
            srcRcpNo.setText(new String(" " + rcpNo + " "));
            srcRcpName.setText(rcpName);
            srcLagNo.setText(new String(" " + lagNo + " "));
            srcLagName.setText(lagName);
            srcMidNo.setText(new String(" " + midNo + " "));
            srcMidName.setText(midName);
            srcItmNo.setText(new String(" " + itmNo + " "));
            srcItmName.setText(itmName);
            dstRoName.setDefault();
            return true;
        }


        //
        //
        //
        public boolean setSendStatus(){
            int idx = 0;
            idx = dstRoName.getSelectedIndex();
            if(0 > idx) return false;
            String sendOp = op_name.getText();
            if(1 > sendOp.length()) return false;

            destMidNo = t6MidTable.getSelectedRow() + 1;
            if(1 > destMidNo) return false;
            String ro     = CZSystem.getRoName();
            String dstRo  = CZSystem.getRoName(idx);
//2003.11.12 syusei
//            if( ro.equals(dstRo) && (midNo == destMidNo)) return false;
            if( ro.equals(dstRo) && (rcpNo == destMidNo)) return false;
//2003.11.12 syusei

//@@            CZSystem.log("CZControlTableCp","T6ItemCopy grp_no[" + grpNo  + "]");
//@@            CZSystem.log("CZControlTableCp","T6ItemCopy ro    [" + ro + "]->[" + dstRo + "]");
//@@            CZSystem.log("CZControlTableCp","T6ItemCopy rec_no[" + rcpNo + "]->[" + rcpNo + "]");
//@@            CZSystem.log("CZControlTableCp","T6ItemCopy mid_no[" + midNo + "]->[" + destMidNo + "]");

            return true;
        }

        //
        // メッセージの表示
        //
        private boolean errorMsg(Object msg[]){
            JOptionPane.showMessageDialog(null,msg,
                    "制御テーブル：Ｔ６項目コピー",
                    JOptionPane.ERROR_MESSAGE);
            return true;
        }

        //
        //
        //
        public class RoNo extends JComboBox {

            RoNo(){
                super();
                try{
                    setName("JComboBox1");
                    setFont(new java.awt.Font("dialog", 0, 16));
                    Vector ro = CZSystem.getRoNameList();
                    if(null == ro){
                        CZSystem.exit(0,"Not Ro No");
                    }
                    for(int i = 0 ; ro.size() > i ; i++){
						String s = CZSystem.RoKetaChg((String)ro.elementAt(i));	// 20050725 炉：表示桁数変更
                        addItem(s);
//                        addItem((String)ro.elementAt(i));
                    }
                    setForeground(java.awt.Color.black);
                    setBackground(java.awt.Color.lightGray);
					setFocusable(false);	/* 2007.08.22 */
                    addActionListener(new ChgRoNo());
//@@                    CZSystem.log("CZControlTableCp","T6ItemCopy new RoNo()");
                }
                catch (Throwable e) {
                    CZSystem.handleException(e);
                }
            }

            //
            //
            //
            public void setDefault(){

                int idx = getSelectedIndex();
                old_idx = idx;      //@@@
//@@                CZSystem.log("CZControlTableCp","T6MidTable setDefault() RoNo[" +
//@@                                        idx + "][" + CZSystem.getRoName(idx) + "]" );
				String s = CZSystem.RoKetaChg(CZSystem.getRoName(idx));	// 20050725 炉：表示桁数変更
//2011.04.14 Y.K レシピタイトルに修正
//                vMidName = null;
//                vMidName  = CZSystem.getCtT6Mid( grpNo,rcpNo,lagNo);
//                if(null != vMidName){
                vRcpName = null;
                vRcpName  = CZSystem.getCtTitle(idx);
                if(null != vRcpName){
//2011.04.14 Y.K レシピタイトルに修正
//                    t6MidTable.setData(grpNo,rcpNo,lagNo,idx);
                    t6MidTable.setData(grpNo,rcpNo);
                }
                if(getHaita(idx)){
                    cp_button.setEnabled(true);
                }
                else {
                    cp_button.setEnabled(false);
                    Object msg[] = {"Ｔ６項目コピー",
                        new String(s + "は"),
//                        new String(CZSystem.getRoName(idx) + "は"),
                        "制御盤、他の端末で修正中です"};
                    errorMsg(msg);
                }
            }

            //
            // 画面消去時の排他開放
            //
            public void releaseHaita(){
                int idx = getSelectedIndex();
//@@                CZSystem.log("CZControlTableCp","T6ItemCopy releaseHaita() 排他[" + idx + "]開放");
                putHaita(idx);
            }

            //
            //
            //
            class ChgRoNo implements ActionListener {
                public void actionPerformed(ActionEvent e){

                    RoNo obj = (RoNo)e.getSource();
                    if(-1 == old_idx){
//@@                        CZSystem.log("CZControlTableCp","T6ItemCopy ChgRoNo 排他１回目");
                    }
                    else {
                        putHaita(old_idx);
//@@                        CZSystem.log("CZControlTableCp","T6ItemCopy ChgRoNo 排他[" +
//@@                                        old_idx + "]->[" + obj.getSelectedIndex() + "]");
                    }
                    obj.setDefault();
                    old_idx = obj.getSelectedIndex();
                }
            }
        }


        /***************************************************
         *
         *       中項目テーブル一覧
         *
         ***************************************************/
        class T6MidTable extends JTable {

            private T6MidTblMdl model = null;

            T6MidTable(){
                super();
                try{
                    setName("T6MidTable");
                    setBounds(0, 0, 200, 200);
                    setAutoCreateColumnsFromModel(true);
                    setSelectionMode(ListSelectionModel.SINGLE_SELECTION);
                    setLocale(new Locale("ja","JP"));
                    setFont(new java.awt.Font("dialog", 0, 12));
                    setRowHeight(17);
                    model = new T6MidTblMdl();
                    setModel(model);
                    DefaultTableColumnModel cmdl = (DefaultTableColumnModel)getColumnModel();
                    TableColumn  colum = null;
                    // レシピーNo　2011.04.14 Y.K 中項目No=>レシピNoに変更
                    colum = cmdl.getColumn(0);
                    colum.setMaxWidth(40);
                    colum.setMinWidth(40);
                    colum.setWidth(40);
                    // タイトル　2011.04.14 Y.K 中項目名称=>レシピNoに変更
                    colum = cmdl.getColumn(1);
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
//@@                CZSystem.log("CZControlTableCP","T6MidTable valueChanged [" +
//@@                            getSelectedRow() + "][" + getSelectedColumn() + "]");
            }

            //
            //
            //
//2011.04.14 Y.K Start
//            public void setData(int gr,int rcp,int lag, int mid){
            public void setData(int gr,int rcp){
//2011.04.14 Y.K End

//@@                CZSystem.log("CZControlTableCP","T6MidTable setData [" +
//@@                                gr + " : " + rcp + " : " + lag+ " : " + mid + "]");
//2011.04.14 Y.K Start
//                CZSystemCtT6MidName midName   = null;
                CZSystemCtTitle rcpName   = null;
//2011.04.14 Y.K End
                String          empty   = new String("");
                for(int i = 0 ; i < 999 ; i++){
                    t6MidTable.setValueAt(empty,i,1);
                }

//2011.04.14 Y.K Start
                for(int i = 0 ; i < vRcpName.size() ; i++){
                    rcpName = (CZSystemCtTitle)vRcpName.elementAt(i);
                     if (gr == rcpName.g_no) {
                        t6MidTable.setValueAt(rcpName.title,rcpName.r_no-1,1);
                    }
                }
//                if ( 0 < rcp ) {
                    setRowSelectionInterval(rcp-1,rcp-1);
                    Rectangle cellRect = getCellRect(rcp-1,0,false);
                    if(cellRect != null){
                        scrollRectToVisible(cellRect);
                    }
//                }
//2011.04.14 Y.K End
                repaint();
            }
        }

        /***************************************************
         *
         *       中項目テーブルクラス：モデル
         *
         ***************************************************/
        public class T6MidTblMdl extends AbstractTableModel {

            final   int TBL_COL     = 2;
            private int TBL_ROW     = 999;
            final String[] names = {" # " ,"タイトル"};
            private Object  data[][];

            //
            T6MidTblMdl(){
                super();
                data = new Object[TBL_ROW][TBL_COL];
                for(int i = 0 ; i < TBL_ROW ; i++){
                    data[i][0] = new Integer(i+1);
                    data[i][1] = new String("");
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
                CZSystem.log("CZControlTableCP","data[" + row + "][" + column + "] = [" + data[row][column] + "][" + aValue + "]" );
            }
        }

        /***************************************************
         *
         *   実行ボタン
         *
         ***************************************************/
        class SendButton implements ActionListener {
            public void actionPerformed(ActionEvent ev){

                if(!setSendStatus()){
                    Object msg[] = {"Ｔ６項目コピー",
                                    "Ｔ６項目、設定者を",
                                    "見直してください"};
                    errorMsg(msg);
                    return;
                 }
                //Send
                String sendOp = op_name.getText();
                int idx = dstRoName.getSelectedIndex();
                String dstRo = CZSystem.getRoName(idx);

                if(!CZSystem.CZControlCopyT6Name(
//2003.11.12 syusei
//                    sendOp,dstRo,grpNo,rcpNo,rcpNo,lagNo,lagNo,midNo,destMidNo,itmNo)){
                    sendOp,dstRo,grpNo,rcpNo,destMidNo,lagNo,lagNo,midNo,midNo,itmNo)){
//2003.11.12 syusei
                    dstRoName.setDefault();
                    Object msg[] = {"Ｔ６項目コピー",
                                    "コピーが失敗しました",
                                    "管理者に問い合わせてください"};
                    errorMsg(msg);
                    return;
                }
                dstRoName.setDefault();
                return ;
            }
        }

        /***************************************************
         *
         *   終了ボタン
         *
         ***************************************************/
        class CancelButton implements ActionListener {
            public void actionPerformed(ActionEvent ev){
                setVisible(false);
                dstRoName.releaseHaita();
            }
        }
    } /* public class T6ItemCopy extends JDialog */
//@@@@
}
