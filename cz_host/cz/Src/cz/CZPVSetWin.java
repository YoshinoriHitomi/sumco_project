package cz;

import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.text.DecimalFormat;
import java.util.Locale;

import javax.swing.ButtonGroup;
import javax.swing.JButton;
import javax.swing.JComboBox;
import javax.swing.JDialog;
import javax.swing.JLabel;
import javax.swing.JOptionPane;
import javax.swing.JRadioButton;
import javax.swing.JTextField;
import javax.swing.plaf.metal.MetalBorders.Flush3DBorder;
import javax.swing.text.AttributeSet;
import javax.swing.text.BadLocationException;
import javax.swing.text.Document;
import javax.swing.text.PlainDocument;

/**********************************************************
 *
 *　　メイン画面：PV表示項目設定
 * @author  (KPK Co.,Ltd.)
 * @version 1.0 (2003/04/01)
 ***********************************************************/
public class CZPVSetWin extends JDialog {
    private static final int X_TIME   = CZSystemPVNamePMM.X_TIME;
    private static final int X_LENGTH = CZSystemPVNamePMM.X_LENGTH;

    private PVBox   box[]    = new PVBox[CZPV.PV_DATA_SET_LENGTH];
    private PVText  text[][] = new PVText[CZPV.PV_DATA_SET_LENGTH][2];
    private ProcBox proc_box = null;

    private JRadioButton minutePad      = null;
    private JRadioButton lengthPad      = null;
    private TimesText    times[]        = new TimesText[2];
    private JLabel       times_lab[]    = new JLabel[2];

    private JButton send_button         = null;
    private JButton cancel_button       = null;
    private TText   op_name             = null;

    //_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
    //　ここからコンストラクタ
    //_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
    CZPVSetWin(){
        super();

        setTitle("ＰＶ表示項目設定");
        setSize(800,570);
        setResizable(false);
        setModal(true);

        getContentPane().setLayout(null);
        // 他基地参照機能    @20131021
        if(CZSystemDefine.REFERENCE_RUN == CZSystem.getRunLevel()){
            getContentPane().setBackground(CZSystemDefine.DEFAULT_REFERENCE_BACKGROUND_COL);
        }else{
            getContentPane().setBackground(CZSystemDefine.DEFAULT_BACKGROUND_COL);
        }

        JLabel  lab = new JLabel("設定者",JLabel.CENTER);
        lab.setBounds(20, 500, 100, 24);
        lab.setLocale(new Locale("ja","JP"));
        lab.setFont(new java.awt.Font("dialog", 0, 16));
        lab.setBorder(new Flush3DBorder());
        lab.setForeground(java.awt.Color.black);
        getContentPane().add(lab);

        op_name = new TText();
        op_name.setBounds(120, 500, 140, 24);
        getContentPane().add(op_name);

        send_button = new JButton("実  行");
        send_button.setBounds(260, 500, 100, 24);
        send_button.setLocale(new Locale("ja","JP"));
        send_button.setFont(new java.awt.Font("dialog", 0, 16));
        send_button.setBorder(new Flush3DBorder());
        send_button.setBackground(java.awt.Color.lightGray);
        send_button.addActionListener(new ChgPVVal());  
        getContentPane().add(send_button);  

        cancel_button = new JButton("終  了");  
        cancel_button.setBounds(650, 500, 100, 24);
        cancel_button.setLocale(new Locale("ja","JP"));
        cancel_button.setFont(new java.awt.Font("dialog", 0, 16));  
        cancel_button.setBorder(new Flush3DBorder());
        cancel_button.setBackground(java.awt.Color.lightGray);  
        cancel_button.addActionListener(new CancelButton());
        getContentPane().add(cancel_button);

        JLabel lab1 = new JLabel("表  示  項  目",JLabel.CENTER);
        lab1.setBounds(20, 20, 500, 24);
        lab1.setLocale(new Locale("ja","JP"));  
        lab1.setFont(new java.awt.Font("dialog", 0, 16));
        lab1.setBorder(new Flush3DBorder());
        lab1.setForeground(java.awt.Color.black);
        getContentPane().add(lab1);

        JLabel min = new JLabel("M  i  n",JLabel.CENTER);
        min.setBounds(540, 20, 100, 24);
        min.setLocale(new Locale("ja","JP"));
        min.setFont(new java.awt.Font("dialog", 0, 16));
        min.setBorder(new Flush3DBorder());
        min.setForeground(java.awt.Color.black);
        getContentPane().add(min);  

        JLabel max = new JLabel("M  a  x",JLabel.CENTER);
        max.setBounds(650, 20, 100, 24);
        max.setLocale(new Locale("ja","JP"));
        max.setFont(new java.awt.Font("dialog", 0, 16));
        max.setBorder(new Flush3DBorder());
        max.setForeground(java.awt.Color.black);
        getContentPane().add(max);  

        for(int i = 0 ;i < CZPV.PV_DATA_SET_LENGTH ;i++){
            box[i] = new PVBox(i);  
            box[i].setBounds(20, 50+(i*30), 500, 24);
            box[i].addActionListener(new ChgPV());  

            getContentPane().add(box[i]);
        } //for end

        for(int i = 0 ;i < CZPV.PV_DATA_SET_LENGTH ;i++){
            text[i][0] = new PVText();  
            text[i][0].setBounds(540, 50+(i*30), 100, 24);  
            getContentPane().add(text[i][0]);

            text[i][1] = new PVText();  
            text[i][1].setBounds(650, 50+(i*30), 100, 24);  
            getContentPane().add(text[i][1]);

        } //for end

        JLabel lab2 = new JLabel("プロセス",JLabel.CENTER);
        lab2.setBounds(20, 370, 100, 24);
        lab2.setLocale(new Locale("ja","JP"));  
        lab2.setFont(new java.awt.Font("dialog", 0, 16));
        lab2.setBorder(new Flush3DBorder());
        lab2.setForeground(java.awt.Color.black);
        getContentPane().add(lab2);

        proc_box = new ProcBox();
        proc_box.setBounds(120, 370, 140, 24);  
        proc_box.addActionListener(new ChgProc());  
        getContentPane().add(proc_box);

        ButtonGroup group = new ButtonGroup();

        minutePad = new JRadioButton("５分   × ");
        minutePad.setMnemonic('M');
        group.add(minutePad);
        minutePad.setSelected(true);
        minutePad.setBorder(new Flush3DBorder());
        minutePad.setBounds(20, 420, 100, 24);  
        getContentPane().add(minutePad);

        lengthPad = new JRadioButton("５mm   × ");
        lengthPad.setMnemonic('L');
        group.add(lengthPad);
        lengthPad.setBorder(new Flush3DBorder());
        lengthPad.setBounds(20, 450, 100, 24);  
        getContentPane().add(lengthPad);

        times[0] = new TimesText();
        times[0].setBounds(120, 420, 40, 24);
        times[0].setText(new String("6"));  
        getContentPane().add(times[0]);

        times[1] = new TimesText();
        times[1].setBounds(120, 450, 40, 24);
        times[1].setText(new String("10"));
        getContentPane().add(times[1]);

        times_lab[0] = new JLabel("30 分",JLabel.CENTER);
        times_lab[0].setBounds(150, 420, 110, 24);  
        times_lab[0].setLocale(new Locale("ja","JP"));  
        times_lab[0].setFont(new java.awt.Font("dialog", 0, 16));
        times_lab[0].setBorder(new Flush3DBorder());
        times_lab[0].setForeground(java.awt.Color.black);
        getContentPane().add(times_lab[0]);

        times_lab[1] = new JLabel("50 mm",JLabel.CENTER);
        times_lab[1].setBounds(150, 450, 110, 24);  
        times_lab[1].setLocale(new Locale("ja","JP"));  
        times_lab[1].setFont(new java.awt.Font("dialog", 0, 16));
        times_lab[1].setBorder(new Flush3DBorder());
        times_lab[1].setForeground(java.awt.Color.black);
        getContentPane().add(times_lab[1]);

        chgProc(CZSystemDefine.READY);  
    }

    //
    //
    //
    public boolean setDefault(){    

//@@        CZSystem.log("CZPVSetWin","setDefault()");  
        int proc = CZSystem.getProcNo();
        chgDefault(proc);

        op_name.setText("");

        return true;
    }

    //
    //
    //
    public boolean chgDefault(int proc){    
//@@        CZSystem.log("CZPVSetWin","chgDefault("+ proc +")");    

        CZSystemPVNamePMM unten = CZSystem.getUnten(proc);  
        // @@@@ null
        if ( null != unten ) {
            proc_box.setSelectedIndex(proc);

            for(int i = 0 ;i < CZPV.PV_DATA_SET_LENGTH ;i++){
                box[i].setSelectedIndex(unten.item[i] - 1);
            }

            switch(unten.x_shubetu){    
            case X_TIME :   minutePad.setSelected(true);
               break;
            case X_LENGTH   :   lengthPad.setSelected(true);
               break;
            default     :   minutePad.setSelected(true);
               break;
            }

            times[0].setText(Integer.toString(unten.x_time));
            times[1].setText(Integer.toString(unten.x_width));  

            for(int i = 0 ;i < CZPV.PV_DATA_SET_LENGTH ;i++){
                text[i][0].setText(Float.toString(unten.min[i]));
                text[i][1].setText(Float.toString(unten.max[i]));
            } //for end

//@@            CZSystem.log("CZPVSetWin setDefault","2");  
            return true;
        } else {
            return false;
        }
    }

    //
    //プロセスに有ったX軸を使用可否を設定           
    //
    public void chgProc(int proc){  

//@@        CZSystem.log("CZPVSetWin","chgProc:Proc No=" + proc );

        switch(proc){
        case CZSystemDefine.READY:
            lengthPad.setEnabled(false);
            break;

        case CZSystemDefine.VAC:
            lengthPad.setEnabled(false);
            break;

        case CZSystemDefine.MELT:
            lengthPad.setEnabled(false);
            break;

        case CZSystemDefine.DIP:
            lengthPad.setEnabled(false);
            break;

        case CZSystemDefine.NECK1:
            lengthPad.setEnabled(true);
            break;

        case CZSystemDefine.NECK2:
            lengthPad.setEnabled(true);
            break;

        case CZSystemDefine.SHOULDER:
            lengthPad.setEnabled(true);
            break;

        case CZSystemDefine.BODY:
            lengthPad.setEnabled(true);
            break;

        case CZSystemDefine.TAIL:
            lengthPad.setEnabled(true);
            break;

        case CZSystemDefine.END:
            lengthPad.setEnabled(false);
            break;

        default :
            lengthPad.setEnabled(false);
            break;
        }

        return ;
    }

    //
    //PVグラフ表示項目に合ったMin,Maxを入れる           
    //
    public boolean chgItem(PVBox box){  
        int no = box.getNo();
        int select = box.getSelectedIndex();

//@@        CZSystem.log("CZPVSetWin","chgItem:Item=" + no + " : " + select + " !!");

        CZSystemPVName pv = CZSystem.getPVName(select);

        String t_min = pvFormat(pv.keta,pv.n_min);  
        String t_max = pvFormat(pv.keta,pv.n_max);  

//@@        CZSystem.log("CZPVSetWin","chgItem:Item=" + t_min + " : " + t_max);

        PVText min = text[no][0];
        PVText max = text[no][1];

        min.setText(t_min);
        max.setText(t_max);
        return true;
    }

    //
    // 表示フォーマット         
    //
    public String pvFormat(int keta, int val){  

        DecimalFormat form = null;  
        StringBuffer  buff = new StringBuffer();

        if(0 >= keta){  
            form = new DecimalFormat("0");  
        }
        else {  
            buff.append("0.");  

            for(int i = 0 ;i < keta ;i++){  
                buff.append("0");
            }
            form = new DecimalFormat(buff.toString());  
        }

        float dat = (float)val / (float)(Math.pow(10,keta));

//@@        CZSystem.log("CZPVSetWin","pvFormat:Item=" + buff.toString() + " : " + val + " : " + dat + " !!");

        return form.format(dat);
    }


    //
    //PVグラフ軸の倍率変更
    //
    public boolean chgTimes(){  
        int val = 0;

        try{    
            String st = times[0].getText();
            val = Integer.valueOf(st).intValue();
            times_lab[0].setText(new String((CZPV.TIME_SCALE_ST * val) + " 分"));
        }
        catch(Exception e){ }

        try{    
            String st = times[1].getText();
            val = Integer.valueOf(st).intValue();
            times_lab[1].setText(new String((CZPV.LENGTH_SCALE_ST * val) + " mm"));
        }
        catch(Exception e){ }
        return true;
    }

    //
    //設定者の入力確認          
    private boolean chkOpe(){

        if(1 > op_name.getText().length()) return false;
        return true;
    }
        
    //
    //設定の確認 表示倍率           
    //
    private boolean chkTimes(){
        int x_time;
        int x_width;

        x_time  = Integer.parseInt(times[0].getText());
        x_width = Integer.parseInt(times[1].getText());

        if(1 > x_time) return false;
        if(1 > x_width) return false;

        return true;
    }

    //
    //設定の確認 Min Max            
    //
    private boolean chkMinMax(){    

        float min;  
        float max;  

        for(int i = 0 ;i < CZPV.PV_DATA_SET_LENGTH ;i++){
            String s1 = text[i][0].getText();
            String s2 = text[i][1].getText();

            min = Float.parseFloat(s1);
            max = Float.parseFloat(s2);

            if(min == max) return false;
            if(min >  max) return false;
        }
        return true;
    }

    //
    //メッセージの表示
    //
    private boolean errorMsg(Object msg[]){
        JOptionPane.showMessageDialog(null,msg,
                       "ＰＶ表示項目入力エラー",    
                       JOptionPane.ERROR_MESSAGE);
            return true;
    }

    //
    //PV表示項目の変更
    //
    class ChgPV implements ActionListener {
        public void actionPerformed(ActionEvent e){

            PVBox box = (PVBox)e.getSource();
            int no = box.getNo();

//@@            CZSystem.log("CZPVSetWin","ChgPV:actionPerformed(" + no + ") !!");  
            if(-1 < no) chgItem(box);
            return ;
        }
    }

    //
    //プロセスの変更    
    //
    class ChgProc implements ActionListener {
        public void actionPerformed(ActionEvent e){
            ProcBox obj = (ProcBox)e.getSource();

            int proc = obj.getSelectedIndex();  

//@@            CZSystem.log("CZPVSetWin","ChgProc:actionPerformed(" + proc + ")" );
            chgProc(proc);  
            chgDefault(proc);
            return ;
        }
    }

    //
    //画面の設定を送信
    //
    class ChgPVVal implements ActionListener {  
        public void actionPerformed(ActionEvent ev){    

            if(!chkOpe()){  
                Object msg[] = {"設定者を入力してくださ！！",   
                                       "",
                                       ""};
                errorMsg(msg);
                return ;
            }

            if(!chkTimes()){    
                Object msg[] = {"表示倍率が１未満です",
                                "",
                                ""};
                errorMsg(msg);
                return ;
            }
            
            if(!chkMinMax()){
                Object msg[] = {"Min-Max の矛盾",   
                                   "",
                                   ""};
                errorMsg(msg);
                return ;
            }

            CZSystemPVNamePMM dat = new CZSystemPVNamePMM();

            dat.p_no = proc_box.getSelectedIndex();

            for(int i = 0 ;i < CZPV.PV_DATA_SET_LENGTH ;i++){
                dat.item[i] = box[i].getSelectedIndex() + 1;
                //dat.item[i] = box[i].getSelectedIndex() ;
            }

            for(int i = 0 ;i < CZPV.PV_DATA_SET_LENGTH ;i++){
                String s1 = text[i][0].getText();
                String s2 = text[i][1].getText();

                dat.min[i] = Float.parseFloat(s1);  
                dat.max[i] = Float.parseFloat(s2);  
            }

            if(lengthPad.isSelected()){
                dat.x_shubetu = dat.X_LENGTH;
            }
            else {  
                dat.x_shubetu = dat.X_TIME;
            }

            dat.x_time  = Integer.parseInt(times[0].getText());
            dat.x_width = Integer.parseInt(times[1].getText());

            CZSystem.CZUntenDefineSend(op_name.getText(),dat);  
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
            //setDefault();
            setVisible(false);
        }
    }

    /*  
    *   
    *   PVグラフMin,Maxを設定するTextField          
    *       数値のみを受け付ける        
    *   
    */  
    public class PVText extends JTextField {    

        PVText(){
            super();
            setFont(new java.awt.Font("dialog", 0, 16));
        }

        protected Document createDefaultModel() {
            return new NumericDocument();
        }

        class NumericDocument extends PlainDocument {
            String validValues = "0123456789.-";

            public void insertString( int offset, String str, AttributeSet a )
                                        throws BadLocationException {

                char[] val = str.toCharArray();
                for (int i = 0;i < val.length;i++) {    
                    if(validValues.indexOf(val[i]) == -1) return;
                }

                super.insertString( offset, str, a );
                return ;
            }
        }
    }

    /*  
    *   
    *   PVグラフ倍率を設定するTextField         
    *       数値のみを受け付ける        
    *   
    */  
    public class TimesText extends JTextField {
        
        TimesText(){    
            super();
            setFont(new java.awt.Font("dialog", 0, 16));
        }

        protected Document createDefaultModel() {
            return new NumericDocument();
        }

        class NumericDocument extends PlainDocument {
            String validValues = "0123456789";  

            public void insertString( int offset, String str, AttributeSet a )
                                                throws BadLocationException {

                if(2 < getLength()) return;
                char[] val = str.toCharArray();
                for (int i = 0;i < val.length;i++) {    
                    if(validValues.indexOf(val[i]) == -1) return;
                }

                super.insertString( offset, str, a );
                chgTimes();
            }

            public void remove(int offs, int len)           
                        throws BadLocationException {
                super.remove(offs,len);
                chgTimes();
            }
        }
    }

    /*  
    *   
    *   PVグラフ項目を選択するComboBox          
    *   
    *   
    */  
    public class PVBox extends JComboBox {  

        private int myNo = -1;  

        PVBox(int no){  
            super();

            myNo = no;  

            setFont(new java.awt.Font("dialog", 0, 16));
            setLocale(new Locale("ja","JP"));
            setForeground(CZPV.PV_COLOR[myNo]);
            setBackground(java.awt.Color.black);
			setFocusable(false);	/* 2007.08.22 */

            StringBuffer buf;
            for(int i = 0 ;i < CZSystemDefine.PV_MAX_LENGTH ;i++)
//@@            for(int i = 0 ;i < CZPV.PV_MAX_LENGTH ;i++)
            {
                CZSystemPVName pv = CZSystem.getPVName(i);  

                if(9  > i){
                    buf = new StringBuffer("  " + (i+1) + " - " + pv.k_name.trim());
                }
                else if(99 > i){    
                    buf = new StringBuffer( " " + (i+1) + " - " + pv.k_name.trim());
                }
                else {  
                    buf = new StringBuffer(     + (i+1) + " - " + pv.k_name.trim());
                }


                int len = buf.length();
                for(int j = 18 ;j > len ;j--){  
                    buf.append(" ");
                }

                addItem(buf + pv.j_name.trim());
            } //for end

        }

        //
        //
        //
        public int getNo(){
            return myNo;
        }
    }

    /*  
    *   
    *   PVグラフのプロセスを選択するComboBox
    *   
    *   
    */  
    public class ProcBox extends JComboBox {    

        ProcBox(){  
            super();

            setFont(new java.awt.Font("dialog", 0, 16));
			setFocusable(false);	/* 2007.08.22 */
            for(int i = 0 ;i < 10 ;i++)         
            {
                addItem(CZSystem.getProcName(i));
            } //fro end
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
                for (int i = 0;i < val.length;i++) {    
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
}
