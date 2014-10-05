package cz;

import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.File;
import java.io.FileOutputStream;
import java.io.BufferedWriter;
import java.io.FileNotFoundException;
import java.io.FileWriter;
import java.io.IOException;
import java.io.IOException;
import java.io.PrintWriter;
import java.util.Locale;
import java.util.Date;
import java.util.Vector;
import java.util.Calendar;
import java.text.SimpleDateFormat;
import java.text.ParseException;

import javax.swing.JTextField;
import javax.swing.BorderFactory;
import javax.swing.JCheckBox;
import javax.swing.JPanel;
import javax.swing.JButton;
import javax.swing.JDialog;
import javax.swing.JLabel;
import javax.swing.JOptionPane;
import javax.swing.plaf.metal.MetalBorders.Flush3DBorder;
import javax.swing.text.AttributeSet;
import javax.swing.text.BadLocationException;
import javax.swing.text.Document;
import javax.swing.text.PlainDocument;
/**
 *   変更履歴出力Window 
 * @author  (SPK Co.,Ltd.)
 * @version 1.0 (2008/10/23)
 * 2008.10.23 H.Nagamine 制御テーブル変更履歴作成
 *
 */

public class CZModify extends JDialog {

    JLabel  label                       = null;
    private JButton     output_btn      = null;
    private JButton     cancel_button   = null;
    private JCheckBox   const_chk_box   = null;
    private JCheckBox   t1_chk_box      = null;
    private JCheckBox   t2_chk_box      = null;
    private JCheckBox   t3_chk_box      = null;
    private JCheckBox   t4_chk_box      = null;
    private JCheckBox   t5_chk_box      = null;
    private JCheckBox   t6_chk_box      = null;
    private StartText   start_txt       = null;
    private EndText     end_txt         = null;

    private JCheckBox   ro_chk_box[]    = new JCheckBox[100];
    private JCheckBox   ro_all_chk_box  = null;

     // 項目
    //
    //
    //
    CZModify(){
        super();

		setSize(660,600);
        setResizable(false);
        setModal(true);
        setTitle("変更履歴出力");

        getContentPane().setLayout(null);
        // 他基地参照機能    @20131021
        if(CZSystemDefine.REFERENCE_RUN == CZSystem.getRunLevel()){
            getContentPane().setBackground(CZSystemDefine.DEFAULT_REFERENCE_BACKGROUND_COL);
        }else{
            getContentPane().setBackground(CZSystemDefine.DEFAULT_BACKGROUND_COL);
        }
        
        // 対象期間
        label = new JLabel("対象期間",JLabel.CENTER);
        label.setBounds(20, 20, 100, 24);
        label.setLocale(new Locale("ja","JP"));
        label.setFont(new java.awt.Font("dialog", 0, 18));
        label.setForeground(java.awt.Color.black);
        label.setBackground(java.awt.Color.lightGray);
        label.setBorder(new Flush3DBorder());
        getContentPane().add(label);

        // 対象期間 開始日付
        start_txt = new StartText();
        start_txt.setBounds(120, 20, 100, 24);
        getContentPane().add(start_txt);

        // ～
        label = new JLabel("～",JLabel.CENTER);
        label.setBounds(220, 20, 50, 24);
        label.setLocale(new Locale("ja","JP"));
        label.setFont(new java.awt.Font("dialog", 0, 18));
        label.setForeground(java.awt.Color.black);
        label.setBorder(new Flush3DBorder());
        getContentPane().add(label);

        // 対象期間 終了日付
        end_txt = new EndText();
        end_txt.setBounds(270, 20, 100, 24);
        getContentPane().add(end_txt);

        //検索対象
        JPanel p = null;
        p = new JPanel();
        p.setBounds( 20, 55, 140, 170);
        p.setLayout(null);
        p.setBorder(BorderFactory.createTitledBorder(new Flush3DBorder(),"検索対象"));
        // 他基地参照機能    @20131021
        if(CZSystemDefine.REFERENCE_RUN == CZSystem.getRunLevel()){
            p.setBackground(CZSystemDefine.DEFAULT_REFERENCE_BACKGROUND_COL);
        }else{
            p.setBackground(CZSystemDefine.DEFAULT_BACKGROUND_COL);
        }
        getContentPane().add(p);

        // 操業定数
        const_chk_box = new JCheckBox("操業定数");
        const_chk_box.setBounds(20, 20, 100, 18);
        const_chk_box.setFont(new java.awt.Font("dialog", 0, 18));
        const_chk_box.setForeground(java.awt.Color.black);
        // 他基地参照機能    @20131021
        if(CZSystemDefine.REFERENCE_RUN == CZSystem.getRunLevel()){
            const_chk_box.setBackground(CZSystemDefine.DEFAULT_REFERENCE_BACKGROUND_COL);
        }else{
            const_chk_box.setBackground(CZSystemDefine.DEFAULT_BACKGROUND_COL);
        }
        const_chk_box.setSelected(true);
        const_chk_box.addActionListener(new Const_chk_box_click());
        p.add(const_chk_box);

        // T1:溶解
        t1_chk_box = new JCheckBox("T1:溶解");
        t1_chk_box.setBounds(20, 40, 100, 18);
        t1_chk_box.setFont(new java.awt.Font("dialog", 0, 18));
        t1_chk_box.setForeground(java.awt.Color.black);
        // 他基地参照機能    @20131021
        if(CZSystemDefine.REFERENCE_RUN == CZSystem.getRunLevel()){
            t1_chk_box.setBackground(CZSystemDefine.DEFAULT_REFERENCE_BACKGROUND_COL);
        }else{
            t1_chk_box.setBackground(CZSystemDefine.DEFAULT_BACKGROUND_COL);
        }
        t1_chk_box.setSelected(false);
        t1_chk_box.addActionListener(new T1_chk_box_click());
        p.add(t1_chk_box);

        // T2:引上
        t2_chk_box = new JCheckBox("T2:引上");
        t2_chk_box.setBounds(20, 60, 100, 18);
        t2_chk_box.setFont(new java.awt.Font("dialog", 0, 18));
        t2_chk_box.setForeground(java.awt.Color.black);
        // 他基地参照機能    @20131021
        if(CZSystemDefine.REFERENCE_RUN == CZSystem.getRunLevel()){
            t2_chk_box.setBackground(CZSystemDefine.DEFAULT_REFERENCE_BACKGROUND_COL);
        }else{
            t2_chk_box.setBackground(CZSystemDefine.DEFAULT_BACKGROUND_COL);
        }
        t2_chk_box.setSelected(false);
        t2_chk_box.addActionListener(new T2_chk_box_click());
        p.add(t2_chk_box);

        // T3:回転
        t3_chk_box = new JCheckBox("T3:回転");
        t3_chk_box.setBounds(20, 80, 100, 18);
        t3_chk_box.setFont(new java.awt.Font("dialog", 0, 18));
        t3_chk_box.setForeground(java.awt.Color.black);
        // 他基地参照機能    @20131021
        if(CZSystemDefine.REFERENCE_RUN == CZSystem.getRunLevel()){
            t3_chk_box.setBackground(CZSystemDefine.DEFAULT_REFERENCE_BACKGROUND_COL);
        }else{
            t3_chk_box.setBackground(CZSystemDefine.DEFAULT_BACKGROUND_COL);
        }
        t3_chk_box.setSelected(false);
        t3_chk_box.addActionListener(new T3_chk_box_click());
        p.add(t3_chk_box);

        // T4:取出
        t4_chk_box = new JCheckBox("T4:取出");
        t4_chk_box.setBounds(20,100, 100, 18);
        t4_chk_box.setFont(new java.awt.Font("dialog", 0, 18));
        t4_chk_box.setForeground(java.awt.Color.black);
        // 他基地参照機能    @20131021
        if(CZSystemDefine.REFERENCE_RUN == CZSystem.getRunLevel()){
            t4_chk_box.setBackground(CZSystemDefine.DEFAULT_REFERENCE_BACKGROUND_COL);
        }else{
            t4_chk_box.setBackground(CZSystemDefine.DEFAULT_BACKGROUND_COL);
        }
        t4_chk_box.setSelected(false);
        t4_chk_box.addActionListener(new T4_chk_box_click());
        p.add(t4_chk_box);

        // T5:圧力
        t5_chk_box = new JCheckBox("T5:圧力");
        t5_chk_box.setBounds(20, 120, 100, 18);
        t5_chk_box.setFont(new java.awt.Font("dialog", 0, 18));
        t5_chk_box.setForeground(java.awt.Color.black);
        // 他基地参照機能    @20131021
        if(CZSystemDefine.REFERENCE_RUN == CZSystem.getRunLevel()){
            t5_chk_box.setBackground(CZSystemDefine.DEFAULT_REFERENCE_BACKGROUND_COL);
        }else{
            t5_chk_box.setBackground(CZSystemDefine.DEFAULT_BACKGROUND_COL);
        }
        t5_chk_box.setSelected(false);
        t5_chk_box.addActionListener(new T5_chk_box_click());
        p.add(t5_chk_box);

        // T6:定数
        t6_chk_box = new JCheckBox("T6:定数");
        t6_chk_box.setBounds(20, 140, 100, 18);
        t6_chk_box.setFont(new java.awt.Font("dialog", 0, 18));
        t6_chk_box.setForeground(java.awt.Color.black);
        // 他基地参照機能    @20131021
        if(CZSystemDefine.REFERENCE_RUN == CZSystem.getRunLevel()){
            t6_chk_box.setBackground(CZSystemDefine.DEFAULT_REFERENCE_BACKGROUND_COL);
        }else{
            t6_chk_box.setBackground(CZSystemDefine.DEFAULT_BACKGROUND_COL);
        }
        t6_chk_box.setSelected(false);
        t6_chk_box.addActionListener(new T6_chk_box_click());
        p.add(t6_chk_box);

        //対象炉番
        JPanel rp = null;
        rp = new JPanel();
        rp.setBounds( 190, 55, 440, 430);
        rp.setLayout(null);
        rp.setBorder(BorderFactory.createTitledBorder(new Flush3DBorder(),"対象炉番"));
        // 他基地参照機能    @20131021
        if(CZSystemDefine.REFERENCE_RUN == CZSystem.getRunLevel()){
            rp.setBackground(CZSystemDefine.DEFAULT_REFERENCE_BACKGROUND_COL);
        }else{
            rp.setBackground(CZSystemDefine.DEFAULT_BACKGROUND_COL);
        }
        getContentPane().add(rp);

        Vector ro = CZSystem.getRoNameList();

        for(int i = 0; i < ro.size(); i++){
            String s = CZSystem.RoKetaChg((String)ro.elementAt(i));
            ro_chk_box[i] = new JCheckBox(s);
            ro_chk_box[i].setBounds( 20+(i/20)*80, 20+(i*20)-(i/20)*400, 80, 18 );
            ro_chk_box[i].setFont(new java.awt.Font("dialog", 0, 18));
            ro_chk_box[i].setForeground(java.awt.Color.black);
            // 他基地参照機能    @20131021
            if(CZSystemDefine.REFERENCE_RUN == CZSystem.getRunLevel()){
                ro_chk_box[i].setBackground(CZSystemDefine.DEFAULT_REFERENCE_BACKGROUND_COL);
            }else{
                ro_chk_box[i].setBackground(CZSystemDefine.DEFAULT_BACKGROUND_COL);
            }
            ro_chk_box[i].addActionListener(new ro_chk_box_click());
            rp.add(ro_chk_box[i]);
        }

        ro_all_chk_box = new JCheckBox("全炉");
        ro_all_chk_box.setBounds( 20 + (ro.size()/20) * 80, 20 + (ro.size()*20) - (ro.size()/20) * 400, 80, 18);
        ro_all_chk_box.setFont(new java.awt.Font("dialog", 1, 18));
        ro_all_chk_box.setForeground(java.awt.Color.black);
        // 他基地参照機能    @20131021
        if(CZSystemDefine.REFERENCE_RUN == CZSystem.getRunLevel()){
            ro_all_chk_box.setBackground(CZSystemDefine.DEFAULT_REFERENCE_BACKGROUND_COL);
        }else{
            ro_all_chk_box.setBackground(CZSystemDefine.DEFAULT_BACKGROUND_COL);
        }
        ro_all_chk_box.addActionListener(new ro_all_chk_box_click());
        rp.add(ro_all_chk_box);

        // 実行
        output_btn = new JButton("実  行");
        output_btn.setBounds(20, 520, 100, 24);
        output_btn.setLocale(new Locale("ja","JP"));
        output_btn.setFont(new java.awt.Font("dialog", 0, 18));
        output_btn.setBorder(new Flush3DBorder());
        output_btn.setForeground(java.awt.Color.black);
        output_btn.addActionListener(new Output_btn_click());
        getContentPane().add(output_btn);

        // cancel 
        cancel_button = new JButton("Cancel");
        cancel_button.setBounds(190, 520, 100, 24);
        cancel_button.setLocale(new Locale("ja","JP"));
        cancel_button.setFont(new java.awt.Font("dialog", 0, 18));
        cancel_button.setBorder(new Flush3DBorder());
        cancel_button.setForeground(java.awt.Color.black);
        cancel_button.addActionListener(new CancelButton());
        getContentPane().add(cancel_button);
    }

    public boolean setDefault(){

        start_txt.setText("");
        end_txt.setText("");

        const_chk_box.setSelected(true);
        t1_chk_box.setSelected(false);
        t2_chk_box.setSelected(false);
        t3_chk_box.setSelected(false);
        t4_chk_box.setSelected(false);
        t5_chk_box.setSelected(false);
        t6_chk_box.setSelected(false);

        Vector ro = CZSystem.getRoNameList();

        for(int i = 0; i < ro.size(); i++){
            ro_chk_box[i].setSelected(false);
        }

        ro_all_chk_box.setSelected(false);

        return true;
    }

    /***************************************************************************
    *
    *       開始日付を入力するTextField
    *
    ***************************************************************************/
    class StartText extends JTextField {

        /**
        *
        */
        StartText(){
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
            String validValues = "0123456789/";

            //
            //
            public void insertString( int offset, String str, AttributeSet a )
                                                    throws BadLocationException {

                if(10 < getLength()) return;
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
    /***************************************************************************
    *
    *       終了を入力するTextField
    *
    ***************************************************************************/
    class EndText extends JTextField {

        /**
        *
        */
        EndText(){
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
            String validValues = "0123456789/";

            //
            //
            public void insertString( int offset, String str, AttributeSet a )
                                                    throws BadLocationException {

                if(10 < getLength()) return;
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
    /***************************************************************************
    *
    *       チェックボックス入力制限
    *
    ***************************************************************************/
    class Const_chk_box_click implements ActionListener {
        public void actionPerformed(ActionEvent ev){
            const_chk_box.setSelected(true);
            t1_chk_box.setSelected(false);
            t2_chk_box.setSelected(false);
            t3_chk_box.setSelected(false);
            t4_chk_box.setSelected(false);
            t5_chk_box.setSelected(false);
            t6_chk_box.setSelected(false);
        }
    }
    class T1_chk_box_click implements ActionListener {
        public void actionPerformed(ActionEvent ev){
            const_chk_box.setSelected(false);
            t1_chk_box.setSelected(true);
            t2_chk_box.setSelected(false);
            t3_chk_box.setSelected(false);
            t4_chk_box.setSelected(false);
            t5_chk_box.setSelected(false);
            t6_chk_box.setSelected(false);
        }
    }
    class T2_chk_box_click implements ActionListener {
        public void actionPerformed(ActionEvent ev){
            const_chk_box.setSelected(false);
            t1_chk_box.setSelected(false);
            t2_chk_box.setSelected(true);
            t3_chk_box.setSelected(false);
            t4_chk_box.setSelected(false);
            t5_chk_box.setSelected(false);
            t6_chk_box.setSelected(false);
        }
    }
    class T3_chk_box_click implements ActionListener {
        public void actionPerformed(ActionEvent ev){
            const_chk_box.setSelected(false);
            t1_chk_box.setSelected(false);
            t2_chk_box.setSelected(false);
            t3_chk_box.setSelected(true);
            t4_chk_box.setSelected(false);
            t5_chk_box.setSelected(false);
            t6_chk_box.setSelected(false);
        }
    }
    class T4_chk_box_click implements ActionListener {
        public void actionPerformed(ActionEvent ev){
            const_chk_box.setSelected(false);
            t1_chk_box.setSelected(false);
            t2_chk_box.setSelected(false);
            t3_chk_box.setSelected(false);
            t4_chk_box.setSelected(true);
            t5_chk_box.setSelected(false);
            t6_chk_box.setSelected(false);
        }
    }
    class T5_chk_box_click implements ActionListener {
        public void actionPerformed(ActionEvent ev){
            const_chk_box.setSelected(false);
            t1_chk_box.setSelected(false);
            t2_chk_box.setSelected(false);
            t3_chk_box.setSelected(false);
            t4_chk_box.setSelected(false);
            t5_chk_box.setSelected(true);
            t6_chk_box.setSelected(false);
        }
    }
    class T6_chk_box_click implements ActionListener {
        public void actionPerformed(ActionEvent ev){
            const_chk_box.setSelected(false);
            t1_chk_box.setSelected(false);
            t2_chk_box.setSelected(false);
            t3_chk_box.setSelected(false);
            t4_chk_box.setSelected(false);
            t5_chk_box.setSelected(false);
            t6_chk_box.setSelected(true);
        }
    }

    class ro_chk_box_click implements ActionListener {
        public void actionPerformed(ActionEvent ev){
            ro_all_chk_box.setSelected(false);
        }
    }

    class ro_all_chk_box_click implements ActionListener {
        public void actionPerformed(ActionEvent ev){
            Vector ro = CZSystem.getRoNameList();

            for(int i = 0; i < ro.size(); i++){
                ro_chk_box[i].setSelected(false);
            }
        }
    }

    /***************************************************************************
    *
    *       実行ボタンクリックイベント
    *
    ***************************************************************************/
    class Output_btn_click implements ActionListener {
        public void actionPerformed(ActionEvent ev){

            Date date1 = null;
            Date date2 = null;

            String chk_box_name = null;
            int tx = 0;

            // 検索条件判定
            if( const_chk_box.isSelected()) {
                chk_box_name = const_chk_box.getText();
            }
            else if( t1_chk_box.isSelected()) {
                chk_box_name = t1_chk_box.getText();
                tx = 1;
            }
            else if( t2_chk_box.isSelected()) {
                chk_box_name = t2_chk_box.getText();
                tx = 2;
            }
            else if( t3_chk_box.isSelected()) {
                chk_box_name = t3_chk_box.getText();
                tx = 3;
            }
            else if( t4_chk_box.isSelected()) {
                chk_box_name = t4_chk_box.getText();
                tx = 4;
            }
            else if( t5_chk_box.isSelected()) {
                chk_box_name = t5_chk_box.getText();
                tx = 5;
            }
            else if( t6_chk_box.isSelected()) {
                chk_box_name = t6_chk_box.getText();
            }

            SimpleDateFormat start_sdf = new SimpleDateFormat("yyyy/MM/dd");
            start_sdf.setLenient(false);

            SimpleDateFormat end_sdf = new SimpleDateFormat("yyyy/MM/dd");
            end_sdf.setLenient(false);

            // 入力した時間が正しいかどうかのチェック
            try {
                date1 = start_sdf.parse(start_txt.getText());
                date2 = end_sdf.parse(end_txt.getText());

                if(date1.compareTo(date2) > 0 ) {
                    Object msg[] = {"入力した日付が正しくありません",
                                        "入力を見直してください！！",
                                        ""};
                    errorMsg(msg);
                    return;
                }
            } catch (ParseException e) {
                Object msg[] = {"入力した日付が正しくありません",
                                    "入力を見直してください！！",
                                    ""};
                errorMsg(msg);
                return;
            }

            // 対象炉番チェック
            Vector ro = CZSystem.getRoNameList();
            int count = 0;
            for(int i = 0; i < ro.size(); i++){
                if((ro_all_chk_box.isSelected() == true) || (ro_chk_box[i].isSelected() == true)){
                    count++;
                }
            }
            if( count == 0 ){
                Object msg[] = {"対象炉番を選択してください！！"};
                errorMsg(msg);
                return;
            }

            CZSystem.log("CZModify","Start day:"+ date1 + " End day:" + date2);

            if( const_chk_box.isSelected()) {
                // 操業定数
                ConstModify(chk_box_name,date1,date2);
            }
            else if( t6_chk_box.isSelected()) {
                // T6
                T6Modify(chk_box_name,date1,date2);
            }
            else {
                // T1,T2,T3,T4,T5
                T1_T5Modify(tx,chk_box_name,date1,date2);
            }
        }
        // 操業定数用
        public void ConstModify(String chk_name,Date date1,Date date2) {

            int TBL_ROW     = 0;
            int TBL_COL     = 9;
            Vector  md_list = null;

            int DataCnt     = 0;

            Object  data[][];


            try {
                SimpleDateFormat output_sdf = new SimpleDateFormat("yyMMdd");

                // CSVデータファイル
                File csv = new File(CZSystem.HISTORY_DATA_PATH,"変更履歴_" + output_sdf.format(date1) + "～" + output_sdf.format(date2) + "_操業定数_" + 
                                    CZSystem.getDateTime("yyMMddHHmm") +  ".csv"); 

                BufferedWriter modify_bw = new BufferedWriter(new FileWriter(csv, false)); // 追記モード

                // ヘッダ部分出力内容
                modify_bw.write("対象期間," + start_txt.getText() + ",～," + end_txt.getText());
                modify_bw.newLine();
                modify_bw.newLine();

                modify_bw.write("対象," + chk_name);
                modify_bw.newLine();
                modify_bw.newLine();
                modify_bw.write("炉番,変更日時,変更者,Bt,変更内容,大項目,中項目,項目Ｎｏ,変更前,変更後");

                Vector ro = CZSystem.getRoNameList();
                for(int ro_idx = 0; ro_idx < ro.size(); ro_idx++){
                    if((ro_all_chk_box.isSelected() == true) || (ro_chk_box[ro_idx].isSelected() == true)){
                        String roName = CZSystem.RoKetaChg((String)ro.elementAt(ro_idx));

                        //検索
                        md_list = CZSystem.getModifyHistoryC(start_txt.getText(),end_txt.getText(),roName);
                        if(null == md_list) {
                            continue;
//                            JOptionPane.showMessageDialog(null,"該当する変更履歴がありませんでした。","出力処理",JOptionPane.INFORMATION_MESSAGE);
//                            return;
                        }

                        CZSystem.log("CZModify","SQL OK!!(const)");

                        // 検索結果格納
                        TBL_ROW = md_list.size();
                        data = new Object[TBL_ROW][TBL_COL];

                        for (int i = 0 ; i < TBL_ROW ; i++) {
                            CZSystemModifyHistoryC md = (CZSystemModifyHistoryC)md_list.elementAt(i);
                            if(null == md) break;
                            data[i][0]  = md.s_time;     // 変更日時
                            data[i][1]  = md.op_name;    // 変更者
                            data[i][2]  = md.batch;      // Bt
                            data[i][3]  = md.message;    // 変更内容
                            data[i][4]  = md.key1;       // 大項目
                            data[i][5]  = md.key2;       // 中項目
                            data[i][6]  = md.key3;       // 項目No
                            data[i][7]  = md.key4;       // 変更前
                            data[i][8]  = md.key5;       // 変更後
                        }
                        CZSystem.log("CZModify","DATA SET OK!!(const)");

                        modify_bw.newLine();

                        for (int i = 0 ; i < TBL_ROW ; i++) {
                            modify_bw.write(roName + "," + data[i][0] + "," + data[i][1] + "," + data[i][2] + "," + data[i][3] + "," + data[i][4] + 
                                            "," + data[i][5] + "," + data[i][6] + "," + data[i][7] + "," + data[i][8]);
                            modify_bw.newLine();
                        }
                        DataCnt++;
                    }
                }
                modify_bw.close();

                if(DataCnt == 0){
                    JOptionPane.showMessageDialog(null,"該当する変更履歴がありませんでした。","出力処理",JOptionPane.INFORMATION_MESSAGE);
                    csv.delete();
                }else{
                    JOptionPane.showMessageDialog(null,"変更履歴を出力しました。","出力処理",JOptionPane.INFORMATION_MESSAGE);
                }

            } catch (FileNotFoundException e) {
              // Fileオブジェクト生成時の例外捕捉
                e.printStackTrace();
            } catch (IOException e) {
              // BufferedWriterオブジェクトのクローズ時の例外捕捉
                e.printStackTrace();
            }
        }
        // T6用
        public void T6Modify(String chk_name,Date date1,Date date2) {

            int TBL_ROW     = 0;
            int TBL_COL     =11;
            Vector  md_list = null;

            int DataCnt     = 0;

            Object  data[][];

            try {
                SimpleDateFormat output_sdf = new SimpleDateFormat("yyMMdd");

                // CSVデータファイル
                File csv = new File(CZSystem.HISTORY_DATA_PATH,"変更履歴_" + output_sdf.format(date1) + "～" + output_sdf.format(date2) + "_T6_" + 
                                    CZSystem.getDateTime("yyMMddHHmm") +  ".csv");

                BufferedWriter modify_bw = new BufferedWriter(new FileWriter(csv, false)); // 追記モード
                // ヘッダ部分出力内容
                modify_bw.write("対象期間," + start_txt.getText() + ",～," + end_txt.getText());
                modify_bw.newLine();
                modify_bw.newLine();

                modify_bw.write("対象," + chk_name);
                modify_bw.newLine();
                modify_bw.newLine();
                modify_bw.write("炉番,変更日時,変更者,Bt,変更内容,レシピNo,大項目,中項目,項目No,変更前,変更後");

                Vector ro = CZSystem.getRoNameList();
                for(int ro_idx = 0; ro_idx < ro.size(); ro_idx++){
                    if((ro_all_chk_box.isSelected() == true) || (ro_chk_box[ro_idx].isSelected() == true)){
                        String roName = CZSystem.RoKetaChg((String)ro.elementAt(ro_idx));

                        //検索
                        md_list = CZSystem.getModifyHistoryT6(start_txt.getText(),end_txt.getText(),roName);

                        if(null == md_list) {
                            continue;
//                            JOptionPane.showMessageDialog(null,"該当する変更履歴がありませんでした。","出力処理",JOptionPane.INFORMATION_MESSAGE);
//                            return;
                        }

                        CZSystem.log("CZModify","SQL OK!!(T6)");

                        // 検索結果格納
                        TBL_ROW = md_list.size();
                        data = new Object[TBL_ROW][TBL_COL];

                        for (int i = 0 ; i < TBL_ROW ; i++) {
                            CZSystemModifyHistoryT6 md = (CZSystemModifyHistoryT6)md_list.elementAt(i);
                            if(null == md) break;
                            data[i][0]   = md.s_time;    // 変更日時
                            data[i][1]   = md.op_name;   // 変更者
                            data[i][2]   = md.batch;     // Bt
                            data[i][3]   = md.message;   // 変更内容
                            data[i][4]   = md.key1;      // テーブルNo
                            data[i][5]   = md.key2;      // レシピNo
                            data[i][6]   = md.key3;      // 大項目
                            data[i][7]   = md.key4;      // 中項目
                            data[i][8]   = md.key5;      // 項目No
                            data[i][9]   = md.key6;      // 変更前
                            data[i][10]  = md.key7;      // 変更後
                        }
                        CZSystem.log("CZModify","DATA SET OK!!(T6)");

                        modify_bw.newLine();

                        for (int i = 0 ; i < TBL_ROW ; i++) {
                            modify_bw.write(roName + "," + data[i][0] + "," + data[i][1] + "," + data[i][2] + "," + data[i][3] + "," + data[i][5] + 
                                            "," + data[i][6] + "," + data[i][7] + "," + data[i][8] + "," + data[i][9] + "," + data[i][10]);
                            modify_bw.newLine();
                        }
                        DataCnt++;
                    }
                }
                modify_bw.close();

                if(DataCnt == 0){
                    JOptionPane.showMessageDialog(null,"該当する変更履歴がありませんでした。","出力処理",JOptionPane.INFORMATION_MESSAGE);
                    csv.delete();
                }else{
                    JOptionPane.showMessageDialog(null,"変更履歴を出力しました。","出力処理",JOptionPane.INFORMATION_MESSAGE);
                }

            } catch (FileNotFoundException e) {
              // Fileオブジェクト生成時の例外捕捉
                e.printStackTrace();
            } catch (IOException e) {
              // BufferedWriterオブジェクトのクローズ時の例外捕捉
                e.printStackTrace();
            }
        }
        // T1～T5用
        public void T1_T5Modify(int tx,String chk_name,Date date1,Date date2) {

            int TBL_ROW1    = 0;
            int TBL_COL1    = 7;
            int TBL_ROW2    = 0;
            int TBL_COL2    = 5;

            int cnt_before  = 0;
            int cnt_after   = 0;
            int plus        = 0;

            Vector  md_list1    = null;
            Vector  md_list2    = null;

            int DataCnt     = 0;

            Object  data1[][];
            Object  data2[][];

            try {
                SimpleDateFormat output_sdf = new SimpleDateFormat("yyMMdd");
                String date_tmp = null;
                String date_sql = null;
                int c_cnt =0;

                // CSVデータファイル
                File csv = new File(CZSystem.HISTORY_DATA_PATH,"変更履歴_" + output_sdf.format(date1) + "～" + output_sdf.format(date2) + "_T" + tx + "_" + 
                                    CZSystem.getDateTime("yyMMddHHmm") +  ".csv");

                BufferedWriter modify_bw = new BufferedWriter(new FileWriter(csv, false)); // 上書き保存
                // ヘッダ部分出力内容
                modify_bw.write("対象期間," + start_txt.getText() + ",～," + end_txt.getText());
                modify_bw.newLine();
                modify_bw.newLine();

                modify_bw.write("対象," + chk_name);
                modify_bw.newLine();
                modify_bw.newLine();
                modify_bw.write("炉番,変更日時,変更者,Bt,変更内容,レシピNo,テーブルNo,項目No,変更前(L軸),変更前(R軸),変更後(L軸),変更後(R軸)");

                Vector ro = CZSystem.getRoNameList();
                for(int ro_idx = 0; ro_idx < ro.size(); ro_idx++){
                    if((ro_all_chk_box.isSelected() == true) || (ro_chk_box[ro_idx].isSelected() == true)){
                        String roName = CZSystem.RoKetaChg((String)ro.elementAt(ro_idx));

                        //検索
                        md_list1 = CZSystem.getModifyHistoryTX1(start_txt.getText(),end_txt.getText(),tx,roName);

                        if(null == md_list1) {
                            continue;
//                            JOptionPane.showMessageDialog(null,"該当する変更履歴がありませんでした。","出力処理",JOptionPane.INFORMATION_MESSAGE);
//                            return;
                        }

                        // 検索結果格納
                        TBL_ROW1 = md_list1.size();
                        data1 = new Object[TBL_ROW1][TBL_COL1];

                        for (int i = 0 ; i < TBL_ROW1 ; i++) {
                            CZSystemModifyHistoryTX1 md1 = (CZSystemModifyHistoryTX1)md_list1.elementAt(i);
                            if(null == md1) break;
                            data1[i][0]   = md1.s_time;    // 変更日時
                            data1[i][1]   = md1.op_name;   // 変更者
                            data1[i][2]   = md1.batch;     // Bt
                            data1[i][3]   = md1.message;   // 変更内容
                            data1[i][4]   = md1.key1;      // テーブルNo
                            data1[i][5]   = md1.key2;      // レシピNo
                            data1[i][6]   = md1.key3;      // テーブルNo

                            // 変更日時をSQL文の条件として使うために一部を抜き出す
                            date_tmp = data1[i][0].toString();
                            date_sql = date_tmp.substring(0,19);

                            //前の検索結果に基づき、検索し列数を取得する
                            // 変更前項目列数
                            cnt_before = CZSystem.getModifyHistoryCnt(0,date_sql,roName);
                            // 変更後項目列数
                            cnt_after = CZSystem.getModifyHistoryCnt(1,date_sql,roName);

                            // 外部結合演算子の有無
                            if(cnt_before == cnt_after) {
                                plus = 0;
                            }
                            else if(cnt_before > cnt_after) {
                                plus = 1;
                            }
                            else if(cnt_before < cnt_after) {
                                plus = 2;
                            }

                            //検索
                            md_list2 = CZSystem.getModifyHistoryTX2(plus,date_sql,roName);

                            if(null == md_list2) {
                                c_cnt++;    //スキップ数をカウント
                                continue;
                            }
                            // 検索結果格納
                            TBL_ROW2 = md_list2.size();
                            data2 = new Object[TBL_ROW2][TBL_COL2];

                            for (int j = 0 ; j < TBL_ROW2 ; j++) {
                                CZSystemModifyHistoryTX2 md2 = (CZSystemModifyHistoryTX2)md_list2.elementAt(j);
                                if(null == md2) break;
                                data2[j][0]   = md2.k_no;          // 項目No
                                // 999999はNULLを示す値  NULLの場合ハイフンにする
                                if( md2.l_val_bf == 999999) {
                                    data2[j][1]   = "-";           // 変更前(L軸) NULL
                                }
                                else {
                                    data2[j][1]   = md2.l_val_bf;  // 変更前(L軸)
                                }

                                if( md2.r_val_bf == 999999) {
                                    data2[j][2]   = "-";           // 変更前(R軸) NULL
                                }
                                else {
                                    data2[j][2]   = md2.r_val_bf;  // 変更前(R軸)
                                }

                                if( md2.l_val_af == 999999) {
                                    data2[j][3]   = "-";           // 変更後(L軸) NULL
                                }
                                else {
                                    data2[j][3]   = md2.l_val_af;  // 変更後(L軸)
                                }

                                if( md2.r_val_af == 999999) {
                                    data2[j][4]   = "-";           // 変更後(R軸) NULL
                                }
                                else {
                                    data2[j][4]   = md2.r_val_af;  // 変更後(R軸)
                                }
                            }

                            CZSystem.log("CZModify","DATA SET OK!!(TX)");

                            modify_bw.newLine();

                            for(int k = 0 ; k < TBL_ROW2 ; k++) {
                                modify_bw.write(roName + "," + data1[i][0] + "," + data1[i][1] + "," + data1[i][2] + "," + data1[i][3] + "," + data1[i][5] + 
                                                "," + data1[i][6] + "," + data2[k][0] + "," + data2[k][1] + "," + data2[k][2] + "," + data2[k][3] + 
                                                "," + data2[k][4]);
                                modify_bw.newLine();
                            }
                            if (i == (TBL_ROW1 - 1)) {
                                  continue;
//                                modify_bw.close();
                            }
                        }
                        if(TBL_ROW1 == c_cnt) {
//                            modify_bw.close();
                            if(csv.delete()) {
                                CZSystem.log("CZModify","R_CT_CHG_HISTORY NO DATA: file delete OK");
                            }
                            else {
                                CZSystem.log("CZModify","R_CT_CHG_HISTORY NO DATA: file delete NG");
                            }
                            continue;
//                            JOptionPane.showMessageDialog(null,"該当する変更履歴がありませんでした。","出力処理",JOptionPane.INFORMATION_MESSAGE);
//                            return;
                        }
                        DataCnt++;
                    }
                }
                modify_bw.close();

                if(DataCnt == 0){
                    JOptionPane.showMessageDialog(null,"該当する変更履歴がありませんでした。","出力処理",JOptionPane.INFORMATION_MESSAGE);
                    csv.delete();
                }else{
                    JOptionPane.showMessageDialog(null,"変更履歴を出力しました。","出力処理",JOptionPane.INFORMATION_MESSAGE);
                }

            } catch (FileNotFoundException e) {
              // Fileオブジェクト生成時の例外捕捉
                e.printStackTrace();
            } catch (IOException e) {
              // BufferedWriterオブジェクトのクローズ時の例外捕捉
                e.printStackTrace();
            }
        }

    }
    // キャンセルボタン
    class CancelButton implements ActionListener {
        public void actionPerformed(ActionEvent ev){

        Vector ro = CZSystem.getRoNameList();

        for(int i = 0; i < ro.size(); i++){
            ro_chk_box[i].setSelected(false);
        }

            setVisible(false);
        }
    }
    // エラーメッセージダイアログ
    private boolean errorMsg(Object msg[]){
        JOptionPane.showMessageDialog(null,msg,
            "入力エラー",
            JOptionPane.ERROR_MESSAGE);
        return true;
    }
}

