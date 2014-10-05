package cz;

import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.util.Locale;

import javax.swing.JButton;
import javax.swing.JDialog;
import javax.swing.plaf.metal.MetalBorders.Flush3DBorder;

/***********************************************************
 *
 *   CCDカメラモニタリング切替えWindow 
 * @author  (KPK Co.,Ltd.)
 * @version 1.0 (2003/04/01)
 *
 ***********************************************************/
public class CZCMSMonitorChg extends JDialog {

    private JButton     send_button   = null;
    private JButton     cancel_button = null;
    
    //
    // ---------- コンストラクタ
    //
    CZCMSMonitorChg(){
        super();
        //画面
        setTitle("CCDモニタ切替え");
        setSize(240,230);
        setResizable(false);
        setModal(false);
        getContentPane().setLayout(null);
        // 他基地参照機能    @20131021
        if(CZSystemDefine.REFERENCE_RUN == CZSystem.getRunLevel()){
            getContentPane().setBackground(CZSystemDefine.DEFAULT_REFERENCE_BACKGROUND_COL);
        }else{
            getContentPane().setBackground(CZSystemDefine.DEFAULT_BACKGROUND_COL);
        }

        //実行ボタン
        send_button = new JButton("実  行");
        send_button.setBounds(20, 160, 100, 24);
        send_button.setLocale(new Locale("ja","JP"));
        send_button.setFont(new java.awt.Font("dialog", 0, 18));
        send_button.setBorder(new Flush3DBorder());
        send_button.setForeground(java.awt.Color.black);
        send_button.addActionListener(new SendButton());
        getContentPane().add(send_button);
        //終了ボタン
        cancel_button = new JButton("終  了");
        cancel_button.setBounds(140, 160, 70, 24);
        cancel_button.setLocale(new Locale("ja","JP"));
        cancel_button.setFont(new java.awt.Font("dialog", 0, 18));
        cancel_button.setBorder(new Flush3DBorder());
        cancel_button.setForeground(java.awt.Color.black);
        cancel_button.addActionListener(new CancelButton());
        getContentPane().add(cancel_button);

    }


    // Method
    // Default値を設定する。
    //
    public boolean setDefault(){

        return true;
    }


    // Class
    // 実行ボタンを処理するActionlistener
    //
    class SendButton implements ActionListener {
        public void actionPerformed(ActionEvent ev){
            //Send
            CZSystem.CZOperateCcdChange(1);
        }
    }


    // Class
    // 終了ボタンを処理するActionlistener
    //
    class CancelButton implements ActionListener {
        public void actionPerformed(ActionEvent ev){
            setDefault();
            setVisible(false);
        }
    }
}
