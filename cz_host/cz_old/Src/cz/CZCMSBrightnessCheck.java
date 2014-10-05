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

/***********************************************************
*   輝度変化チェックWindow
* @author
* @version 1.0
*
************************************************************/
public class CZCMSBrightnessCheck extends JDialog {

    JLabel  label                       = null;
    private JButton     output_btn      = null;
    private JButton     cancel_button   = null;
    private JCheckBox   neck_chk_box    = null;
    private JCheckBox   shold_chk_box   = null;
    private JCheckBox   body_chk_box    = null;
    private JCheckBox   all_chk_box     = null;
    private StartText   start_txt       = null;
    private EndText     end_txt         = null;

    private JCheckBox   ro_chk_box[]    = new JCheckBox[100];
    private JCheckBox   ro_all_chk_box  = null;

    /*******************************************************
    *
    *   コンストラクタ
    *
    ********************************************************/
    CZCMSBrightnessCheck(){
        super();

        setSize(660,600);
        setResizable(false);
        setModal(true);
        setTitle("輝度変化チェック");

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

        //出力対象プロセス
        JPanel p = null;
        p = new JPanel();
        p.setBounds( 20, 55, 140, 90);
        p.setLayout(null);
        p.setBorder(BorderFactory.createTitledBorder(new Flush3DBorder(),"出力対象プロセス"));
        // 他基地参照機能    @20131021
        if(CZSystemDefine.REFERENCE_RUN == CZSystem.getRunLevel()){
            p.setBackground(CZSystemDefine.DEFAULT_REFERENCE_BACKGROUND_COL);
        }else{
            p.setBackground(CZSystemDefine.DEFAULT_BACKGROUND_COL);
        }
        getContentPane().add(p);

        // NECK
        neck_chk_box = new JCheckBox("NECK");
        neck_chk_box.setBounds(20, 20, 100, 18);
        neck_chk_box.setFont(new java.awt.Font("dialog", 0, 18));
        neck_chk_box.setForeground(java.awt.Color.black);
        // 他基地参照機能    @20131021
        if(CZSystemDefine.REFERENCE_RUN == CZSystem.getRunLevel()){
            neck_chk_box.setBackground(CZSystemDefine.DEFAULT_REFERENCE_BACKGROUND_COL);
        }else{
            neck_chk_box.setBackground(CZSystemDefine.DEFAULT_BACKGROUND_COL);
        }
        neck_chk_box.setSelected(false);
//        neck_chk_box.addActionListener(new neck_chk_box_click());
        p.add(neck_chk_box);

        // SHOLD
        shold_chk_box = new JCheckBox("SHOLD");
        shold_chk_box.setBounds(20, 40, 100, 18);
        shold_chk_box.setFont(new java.awt.Font("dialog", 0, 18));
        shold_chk_box.setForeground(java.awt.Color.black);
        // 他基地参照機能    @20131021
        if(CZSystemDefine.REFERENCE_RUN == CZSystem.getRunLevel()){
            shold_chk_box.setBackground(CZSystemDefine.DEFAULT_REFERENCE_BACKGROUND_COL);
        }else{
            shold_chk_box.setBackground(CZSystemDefine.DEFAULT_BACKGROUND_COL);
        }
        shold_chk_box.setSelected(false);
//        shold_chk_box.addActionListener(new shold_chk_box_click());
        p.add(shold_chk_box);

        // BODY
        body_chk_box = new JCheckBox("BODY");
        body_chk_box.setBounds(20, 60, 100, 18);
        body_chk_box.setFont(new java.awt.Font("dialog", 0, 18));
        body_chk_box.setForeground(java.awt.Color.black);
        // 他基地参照機能    @20131021
        if(CZSystemDefine.REFERENCE_RUN == CZSystem.getRunLevel()){
            body_chk_box.setBackground(CZSystemDefine.DEFAULT_REFERENCE_BACKGROUND_COL);
        }else{
            body_chk_box.setBackground(CZSystemDefine.DEFAULT_BACKGROUND_COL);
        }
        body_chk_box.setSelected(false);
//        body_chk_box.addActionListener(new body_chk_box_click());
        p.add(body_chk_box);

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
            int neck_prc  = 0;
            int shold_prc = 0;
            int body_prc  = 0;
            boolean o_flg = false;

            // 出力対象期間チェック
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

            // 出力対象プロセスチェック
            if( neck_chk_box.isSelected()) {
                chk_box_name = neck_chk_box.getText();
                neck_prc  = 4;
                o_flg = true;
            }

            if( shold_chk_box.isSelected()) {
                chk_box_name = shold_chk_box.getText();
                shold_prc = 6;
                o_flg = true;
            }

            if( body_chk_box.isSelected()) {
                chk_box_name = body_chk_box.getText();
                body_prc  = 7;
                o_flg = true;
            }

            if(o_flg == false){
                Object msg[] = {"出力対象プロセスが選択されていません",
                                "プロセスを選択してください！！",
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

            CZSystem.log("CZCMSBrightnessCheck","Start day:"+ date1 + " End day:" + date2);

            // プロセス毎処理
            if( neck_chk_box.isSelected()) {
                // NECK
                BrightnessData_OutPut(neck_prc,chk_box_name,date1,date2);
                BrightnessData2_OutPut(neck_prc,chk_box_name,date1,date2);
            }

            if( shold_chk_box.isSelected()) {
                // SHOLD
                BrightnessData_OutPut(shold_prc,chk_box_name,date1,date2);
                BrightnessData2_OutPut(shold_prc,chk_box_name,date1,date2);
            }

            if( body_chk_box.isSelected()) {
                // BODY
                BrightnessData_OutPut(body_prc,chk_box_name,date1,date2);
            }

            JOptionPane.showMessageDialog(null,"データを出力しました。","出力処理",JOptionPane.INFORMATION_MESSAGE);

        }

        ////////////////////////////////////////////////////////////////////////
        //
        // 輝度変化チェックデータ CSVファイル出力（最大輝度比較）
        //
        ////////////////////////////////////////////////////////////////////////
        public void BrightnessData_OutPut(int prc,String chk_name,Date date1,Date date2) {

            String  pName       = null;

            int     TBL_ROW     = 0;
            int     TBL_COL     = 27;

            int     cnt_before  = 0;
            int     cnt_after   = 0;
            int     plus        = 0;

            Vector  bcd_list    = null;

            int     DataCnt     = 0;

            Object  data[][];

            int     val;
            String  a           = null;
            char    c[];
            int     size        = 0;
            int     len;
            int     word        = 2;

            try {
                SimpleDateFormat output_sdf = new SimpleDateFormat("yyMMdd");
                String date_tmp = null;
                String date_sql = null;
                int c_cnt = 0;

                // CSVデータファイル
                switch (prc) {
                    case 4:  // NECK
                        pName = "NECK";
                        break;

                    case 6:  // SHOLD
                        pName = "SHOLD";
                        break;

                    case 7:  // BODY
                        pName = "BODY";
                        break;
                }

                File csv = new File(CZSystem.KIDO_DATA_PATH,"(最大輝度比較)" + pName +"_" + output_sdf.format(date1) + "～" + output_sdf.format(date2) + "_" + 
                                    CZSystem.getDateTime("yyMMddHHmm") +  ".csv");

                BufferedWriter modify_bw = new BufferedWriter(new FileWriter(csv, false)); // 上書き保存
                // ヘッダ部分出力内容
                modify_bw.write("対象期間," + start_txt.getText() + ",～," + end_txt.getText());
                modify_bw.newLine();
                modify_bw.newLine();

                modify_bw.write("対象プロセス," + pName);
                modify_bw.newLine();
                modify_bw.newLine();

                // CSVファイル ヘッダー部
                switch (prc) {
                    case 4: // NECK
                    case 6: // SHOLD
                        modify_bw.write("炉番,採取日時,チャージ,GAP,バッチNo,プロセスNo,NS:最大輝度平均,比較対象バッチNo," + 
                                        "(比較対象)NS:最大輝度平均,NS:最大輝度判定閾値,(閾値判定対象値)NS:最大輝度平均,輝度データ");
                    break;

                    case 7: // BODY
                        modify_bw.write("炉番,採取日時,チャージ,GAP,バッチNo,プロセスNo,B:(左)最大輝度平均,B:(右)最大輝度平均,B:片ピーク," +
                                        "比較対象バッチNo,(比較対象)B:(左)最大輝度平均,(比較対象)B:(右)最大輝度平均," +
                                        "B:最大輝度判定閾値,B:片ピーク判定閾値,(閾値判定対象値)B:(左)最大輝度平均,(閾値判定対象値)B:(右)最大輝度平均," +
                                        "輝度データ");
                    break;
                }
                modify_bw.newLine();

                Vector ro = CZSystem.getRoNameList();
                for(int ro_idx = 0; ro_idx < ro.size(); ro_idx++){
                    if((ro_all_chk_box.isSelected() == true) || (ro_chk_box[ro_idx].isSelected() == true)){
                        String roName = CZSystem.RoKetaChg((String)ro.elementAt(ro_idx));

                        //検索
                        bcd_list = CZSystem.getBrightnessCsvData(start_txt.getText(),end_txt.getText(),prc,roName);

                        if(null == bcd_list) {
                            continue;
//                            JOptionPane.showMessageDialog(null,"該当する出力データがありませんでした。","出力処理",JOptionPane.INFORMATION_MESSAGE)
//                            return;
                        }

                        // 検索結果格納
                        TBL_ROW = bcd_list.size();
                        data = new Object[TBL_ROW][TBL_COL];

                        for (int i = 0 ; i < TBL_ROW ; i++) {
                            CZSystemBrightnessData bcd = (CZSystemBrightnessData)bcd_list.elementAt(i);
                            if(null == bcd) break;
                            data[i][0]   = bcd.s_time;             // 採取日時
                            data[i][1]   = bcd.batch;              // バッチNo
                            data[i][2]   = bcd.p_no;               // プロセスNo
                            data[i][3]   = bcd.charge;             // チャージ量
                            data[i][4]   = bcd.gap;                // GAP
                            data[i][5]   = bcd.max_b_ave;          // NS:最大輝度平均
                            data[i][6]   = bcd.range_b_ave;        // NS:指定区間輝度平均
                            data[i][7]   = bcd.max_b_judge;        // NS:最大輝度判定閾値
                            data[i][8]   = bcd.range_b_judge;      // NS:指定区間輝度判定閾値
                            data[i][9]   = bcd.x_review;           // NS:評価X座標
                            data[i][10]  = bcd.review_range;       // NS:評価範囲
                            data[i][11]  = bcd.body_l_max_b_ave;   // B:(左)最大輝度平均
                            data[i][12]  = bcd.body_r_max_b_ave;   // B:(右)最大輝度平均
                            data[i][13]  = bcd.body_max_b_range;   // B:最大輝度判定閾値
                            data[i][14]  = bcd.body_peek;          // B:片ピーク
                            data[i][15]  = bcd.body_peek_judge;    // B:片ピーク判定閾値
                            data[i][16]  = bcd.len;                // データ数
                            data[i][17]  = bcd.data;               // データ
                            data[i][18]  = bcd.c_batch;            // 比較対象バッチNo
                            data[i][19]  = bcd.c_max_b_ave;        // (比較対象)NS:最大輝度平均
                            data[i][20]  = bcd.c_range_b_ave;      // (比較対象)NS:指定区間輝度平均
                            data[i][21]  = bcd.t_max_b_judge;      // (閾値判定対象値)NS:最大輝度平均
                            data[i][22]  = bcd.t_range_b_judge;    // (閾値判定対象値)NS:指定区間輝度平均
                            data[i][23]  = bcd.c_body_l_max_b_ave; // (比較対象)B:(左)最大輝度平均
                            data[i][24]  = bcd.c_body_r_max_b_ave; // (比較対象)B:(右)最大輝度平均
                            data[i][25]  = bcd.t_body_l_max_b_ave; // (閾値判定対象値)B:(左)最大輝度平均
                            data[i][26]  = bcd.t_body_r_max_b_ave; // (閾値判定対象値)B:(右)最大輝度平均

                            CZSystem.log("CZCMSBrightnessCheck","DATA SET OK!!(NS)");


			                size = bcd.len;
			                len  = size * word;

			                if (null == data[i][17]) continue;  //@@
			                c = data[i][17].toString().toCharArray();

                            // CSVファイル データ部
                            switch (prc) {
                                case 4: // NECK
                                case 6: // SHOLD
                                    modify_bw.write(roName + "," + data[i][0] + "," + data[i][3] + "," + data[i][4] + "," + data[i][1] + "," + data[i][2] + 
                                                    "," + data[i][5] + "," + data[i][18] + "," + data[i][19] + "," + data[i][7] + "," + data[i][21]);

					                for(int ii = 0 ; ii < len ; ii+=word){
					                    a = new String(c,ii,word);
					                    val = Integer.parseInt(a,16);
                                        modify_bw.write("," + val);
					                }

                                break;

                                case 7: // BODY
                                    modify_bw.write(roName + "," + data[i][0] + "," + data[i][3] + "," + data[i][4] + "," + data[i][1] + "," + data[i][2] + 
                                                    "," + data[i][11] + "," + data[i][12] + "," + data[i][14] + "," + data[i][18] + "," + data[i][23] + 
                                                    "," + data[i][24] + "," + data[i][13] + "," + data[i][15] + "," + data[i][25] + "," + data[i][26]);

					                for(int ii = 0 ; ii < len ; ii+=word){
					                    a = new String(c,ii,word);
					                    val = Integer.parseInt(a,16);
                                        modify_bw.write("," + val);
					                }

                                break;
                            }
                            modify_bw.newLine();

                            if (i == (TBL_ROW - 1)) {
                                  continue;
//                                modify_bw.close();
                            }
                        }
                        if(TBL_ROW == c_cnt) {
//                            modify_bw.close();
                            if(csv.delete()) {
                                CZSystem.log("CZCMSBrightnessCheck","R_BRIGHTNESS_CHANGE NO DATA: file delete OK");
                            }
                            else {
                                CZSystem.log("CZCMSBrightnessCheck","R_BRIGHTNESS_CHANGE NO DATA: file delete NG");
                            }
                            continue;
//                            JOptionPane.showMessageDialog(null,"該当する出力データがありませんでした。","出力処理",JOptionPane.INFORMATION_MESSAGE);
//                            return;
                        }
                        DataCnt++;
                    }
                }
                modify_bw.close();

                if(DataCnt == 0){
                    JOptionPane.showMessageDialog(null,"(最大輝度比較):" + pName + " 該当する出力データがありませんでした。","出力処理",JOptionPane.INFORMATION_MESSAGE);
                    csv.delete();
                }else{
                    // JOptionPane.showMessageDialog(null,"データを出力しました。","出力処理",JOptionPane.INFORMATION_MESSAGE);
                }

            } catch (FileNotFoundException e) {
              // Fileオブジェクト生成時の例外捕捉
                e.printStackTrace();
            } catch (IOException e) {
              // BufferedWriterオブジェクトのクローズ時の例外捕捉
                e.printStackTrace();
            }

        }

        ////////////////////////////////////////////////////////////////////////
        //
        // 輝度変化チェックデータ CSVファイル出力（平均データ）
        //
        ////////////////////////////////////////////////////////////////////////
        public void BrightnessData2_OutPut(int prc,String chk_name,Date date1,Date date2) {

            String  pName       = null;

            int     TBL_ROW     = 0;
            int     TBL_COL     = 27;

            int     cnt_before  = 0;
            int     cnt_after   = 0;
            int     plus        = 0;

            Vector  bcd_list    = null;

            int     DataCnt     = 0;

            Object  data[][];

            int     val;
            String  a           = null;
            char    c[];
            int     size        = 0;
            int     len;
            int     word        = 2;

            try {
                SimpleDateFormat output_sdf = new SimpleDateFormat("yyMMdd");
                String date_tmp = null;
                String date_sql = null;
                int c_cnt = 0;

                // CSVデータファイル
                switch (prc) {
                    case 4:  // NECK
                        pName = "NECK";
                        break;

                    case 6:  // SHOLD
                        pName = "SHOLD";
                        break;

                }
                File csv = new File(CZSystem.KIDO_DATA_PATH,"(平均データ)" + pName + "_" + output_sdf.format(date1) + "～" + output_sdf.format(date2) + "_" + 
                                    CZSystem.getDateTime("yyMMddHHmm") +  ".csv");


                BufferedWriter modify_bw = new BufferedWriter(new FileWriter(csv, false)); // 上書き保存
                // ヘッダ部分出力内容
                modify_bw.write("対象期間," + start_txt.getText() + ",～," + end_txt.getText());
                modify_bw.newLine();
                modify_bw.newLine();

                modify_bw.write("対象プロセス," + pName);
                modify_bw.newLine();
                modify_bw.newLine();

                // CSVファイル ヘッダー部
                modify_bw.write("炉番,採取日時,チャージ,GAP,バッチNo,プロセスNo,NS:指定区間輝度平均,比較対象バッチNo," + 
                                "(比較対象)NS:指定区間輝度平均,NS:指定区間輝度判定閾値,(閾値判定対象値)NS:指定区間輝度平均," +
                                "NS:評価X座標,NS:評価範囲,輝度データ");

                modify_bw.newLine();

                Vector ro = CZSystem.getRoNameList();
                for(int ro_idx = 0; ro_idx < ro.size(); ro_idx++){
                    if((ro_all_chk_box.isSelected() == true) || (ro_chk_box[ro_idx].isSelected() == true)){
                        String roName = CZSystem.RoKetaChg((String)ro.elementAt(ro_idx));

                        //検索
                        bcd_list = CZSystem.getBrightnessCsvData(start_txt.getText(),end_txt.getText(),prc,roName);

                        if(null == bcd_list) {
                            continue;
//                            JOptionPane.showMessageDialog(null,"該当する出力データがありませんでした。","出力処理",JOptionPane.INFORMATION_MESSAGE)
//                            return;
                        }

                        // 検索結果格納
                        TBL_ROW = bcd_list.size();
                        data = new Object[TBL_ROW][TBL_COL];

                        for (int i = 0 ; i < TBL_ROW ; i++) {
                            CZSystemBrightnessData bcd = (CZSystemBrightnessData)bcd_list.elementAt(i);
                            if(null == bcd) break;
                            data[i][0]   = bcd.s_time;             // 採取日時
                            data[i][1]   = bcd.batch;              // バッチNo
                            data[i][2]   = bcd.p_no;               // プロセスNo
                            data[i][3]   = bcd.charge;             // チャージ量
                            data[i][4]   = bcd.gap;                // GAP
                            data[i][5]   = bcd.max_b_ave;          // NS:最大輝度平均
                            data[i][6]   = bcd.range_b_ave;        // NS:指定区間輝度平均
                            data[i][7]   = bcd.max_b_judge;        // NS:最大輝度判定閾値
                            data[i][8]   = bcd.range_b_judge;      // NS:指定区間輝度判定閾値
                            data[i][9]   = bcd.x_review;           // NS:評価X座標
                            data[i][10]  = bcd.review_range;       // NS:評価範囲
                            data[i][11]  = bcd.body_l_max_b_ave;   // B:(左)最大輝度平均
                            data[i][12]  = bcd.body_r_max_b_ave;   // B:(右)最大輝度平均
                            data[i][13]  = bcd.body_max_b_range;   // B:最大輝度判定閾値
                            data[i][14]  = bcd.body_peek;          // B:片ピーク
                            data[i][15]  = bcd.body_peek_judge;    // B:片ピーク判定閾値
                            data[i][16]  = bcd.len;                // データ数
                            data[i][17]  = bcd.data;               // データ
                            data[i][18]  = bcd.c_batch;            // 比較対象バッチNo
                            data[i][19]  = bcd.c_max_b_ave;        // (比較対象)NS:最大輝度平均
                            data[i][20]  = bcd.c_range_b_ave;      // (比較対象)NS:指定区間輝度平均
                            data[i][21]  = bcd.t_max_b_judge;      // (閾値判定対象値)NS:最大輝度平均
                            data[i][22]  = bcd.t_range_b_judge;    // (閾値判定対象値)NS:指定区間輝度平均
                            data[i][23]  = bcd.c_body_l_max_b_ave; // (比較対象)B:(左)最大輝度平均
                            data[i][24]  = bcd.c_body_r_max_b_ave; // (比較対象)B:(右)最大輝度平均
                            data[i][25]  = bcd.t_body_l_max_b_ave; // (閾値判定対象値)B:(左)最大輝度平均
                            data[i][26]  = bcd.t_body_r_max_b_ave; // (閾値判定対象値)B:(右)最大輝度平均

                            CZSystem.log("CZCMSBrightnessCheck","DATA SET OK!!(NS)");

			                size = bcd.len;
			                len  = size * word;

			                if (null == data[i][17]) continue;  //@@
			                c = data[i][17].toString().toCharArray();

                            // CSVファイル データ部
                            modify_bw.write(roName + "," + data[i][0] + "," + data[i][3] + "," + data[i][4] + "," + data[i][1] + "," + data[i][2] + 
                                            "," + data[i][6] + "," + data[i][18] + "," + data[i][20] + "," + data[i][8] + "," + data[i][22] + 
                                            "," + data[i][9] + "," + data[i][10]);

			                for(int ii = 0 ; ii < len ; ii+=word){
			                    a = new String(c,ii,word);
			                    val = Integer.parseInt(a,16);
                                modify_bw.write("," + val);
			                }

                            modify_bw.newLine();

                            if (i == (TBL_ROW - 1)) {
                                  continue;
//                                modify_bw.close();
                            }
                        }
                        if(TBL_ROW == c_cnt) {
//                            modify_bw.close();
                            if(csv.delete()) {
                                CZSystem.log("CZCMSBrightnessCheck","R_BRIGHTNESS_CHANGE NO DATA: file delete OK");
                            }
                            else {
                                CZSystem.log("CZCMSBrightnessCheck","R_BRIGHTNESS_CHANGE NO DATA: file delete NG");
                            }
                            continue;
//                            JOptionPane.showMessageDialog(null,"該当する出力データがありませんでした。","出力処理",JOptionPane.INFORMATION_MESSAGE);
//                            return;
                        }
                        DataCnt++;
                    }
                }
                modify_bw.close();

                if(DataCnt == 0){
                    JOptionPane.showMessageDialog(null,"(平均データ):" + pName + " 該当する出力データがありませんでした。","出力処理",JOptionPane.INFORMATION_MESSAGE);
                    csv.delete();
                }else{
                    // JOptionPane.showMessageDialog(null,"データを出力しました。","出力処理",JOptionPane.INFORMATION_MESSAGE);
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

    /***************************************************************************
    *
    *       キャンセルボタンクリックイベント
    *
    ***************************************************************************/
    class CancelButton implements ActionListener {
        public void actionPerformed(ActionEvent ev){

        Vector ro = CZSystem.getRoNameList();

        for(int i = 0; i < ro.size(); i++){
            ro_chk_box[i].setSelected(false);
        }

            setVisible(false);
        }
    }

    ////////////////////////////////////////////////////////
    //
    // エラーメッセージダイアログ
    //
    ////////////////////////////////////////////////////////
    private boolean errorMsg(Object msg[]){
        JOptionPane.showMessageDialog(null,msg,
            "入力エラー",
            JOptionPane.ERROR_MESSAGE);
        return true;
    }
}

