package cz;

import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.util.Locale;

import javax.swing.JButton;
import javax.swing.JDialog;
import javax.swing.plaf.metal.MetalBorders.Flush3DBorder;

/***********************************************************
 *
 *   CCD�J�������j�^�����O�ؑւ�Window 
 * @author  (KPK Co.,Ltd.)
 * @version 1.0 (2003/04/01)
 *
 ***********************************************************/
public class CZCMSMonitorChg extends JDialog {

    private JButton     send_button   = null;
    private JButton     cancel_button = null;
    
    //
    // ---------- �R���X�g���N�^
    //
    CZCMSMonitorChg(){
        super();
        //���
        setTitle("CCD���j�^�ؑւ�");
        setSize(240,230);
        setResizable(false);
        setModal(false);
        getContentPane().setLayout(null);
        // ����n�Q�Ƌ@�\    @20131021
        if(CZSystemDefine.REFERENCE_RUN == CZSystem.getRunLevel()){
            getContentPane().setBackground(CZSystemDefine.DEFAULT_REFERENCE_BACKGROUND_COL);
        }else{
            getContentPane().setBackground(CZSystemDefine.DEFAULT_BACKGROUND_COL);
        }

        //���s�{�^��
        send_button = new JButton("��  �s");
        send_button.setBounds(20, 160, 100, 24);
        send_button.setLocale(new Locale("ja","JP"));
        send_button.setFont(new java.awt.Font("dialog", 0, 18));
        send_button.setBorder(new Flush3DBorder());
        send_button.setForeground(java.awt.Color.black);
        send_button.addActionListener(new SendButton());
        getContentPane().add(send_button);
        //�I���{�^��
        cancel_button = new JButton("�I  ��");
        cancel_button.setBounds(140, 160, 70, 24);
        cancel_button.setLocale(new Locale("ja","JP"));
        cancel_button.setFont(new java.awt.Font("dialog", 0, 18));
        cancel_button.setBorder(new Flush3DBorder());
        cancel_button.setForeground(java.awt.Color.black);
        cancel_button.addActionListener(new CancelButton());
        getContentPane().add(cancel_button);

    }


    // Method
    // Default�l��ݒ肷��B
    //
    public boolean setDefault(){

        return true;
    }


    // Class
    // ���s�{�^������������Actionlistener
    //
    class SendButton implements ActionListener {
        public void actionPerformed(ActionEvent ev){
            //Send
            CZSystem.CZOperateCcdChange(1);
        }
    }


    // Class
    // �I���{�^������������Actionlistener
    //
    class CancelButton implements ActionListener {
        public void actionPerformed(ActionEvent ev){
            setDefault();
            setVisible(false);
        }
    }
}
