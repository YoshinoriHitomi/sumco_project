package cz;

import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.util.Locale;

import javax.swing.JButton;
import javax.swing.JDialog;
import javax.swing.JLabel;
import javax.swing.JOptionPane;
import javax.swing.JScrollPane;
import javax.swing.JTextField;
import javax.swing.JTree;
import javax.swing.event.TreeSelectionEvent;
import javax.swing.event.TreeSelectionListener;
import javax.swing.plaf.metal.MetalBorders.Flush3DBorder;
import javax.swing.text.AttributeSet;
import javax.swing.text.BadLocationException;
import javax.swing.text.Document;
import javax.swing.text.PlainDocument;
import javax.swing.tree.DefaultMutableTreeNode;
import javax.swing.tree.DefaultTreeSelectionModel;
import javax.swing.tree.TreePath;
import javax.swing.tree.TreeSelectionModel;

/**
 *   ���ƒ萔�R�s�[�pWindow 
 * @author  (KPK Co.,Ltd.)
 * @version 1.0 (2003/04/01)
 *
 */

public class CZOperationTableCp extends JDialog {


    private final int NO_COPY     = -1;
    private final int ALL_COPY    =  1;
    private final int LARGE_COPY  =  2;
    private final int MIDDLE_COPY =  3;

    private int copy_mode         =  NO_COPY;

    private JLabel  ro_lab   = null;
    private JLabel  copy_lab = null;

    private RoText      ro_name  = null;

    private JButton     send_button   = null;
    private JButton     cancel_button = null;

    private TText       op_name  = null;

    private DefaultMutableTreeNode  top     = null; 
    private JTree           tree        = null; 
    private JScrollPane     treepanel   = null;


    private int current_lag = -1;
    private int current_mid = -1;

    //
    //
    //
    CZOperationTableCp(){
        super();

        setTitle("���ƒ萔�R�s�[");
        setSize(500,480);
        setResizable(false);
        setModal(true);

        getContentPane().setLayout(null);
        // ����n�Q�Ƌ@�\    @20131021
        if(CZSystemDefine.REFERENCE_RUN == CZSystem.getRunLevel()){
            getContentPane().setBackground(CZSystemDefine.DEFAULT_REFERENCE_BACKGROUND_COL);
        }else{
            getContentPane().setBackground(CZSystemDefine.DEFAULT_BACKGROUND_COL);
        }

        JLabel  lab = new JLabel("�R�s�[��",JLabel.CENTER);
        lab.setBounds(20, 20, 100, 30);
        lab.setLocale(new Locale("ja","JP"));
        lab.setFont(new java.awt.Font("dialog", 0, 18));
        lab.setBorder(new Flush3DBorder());
        lab.setForeground(java.awt.Color.black);
        getContentPane().add(lab);

        ro_lab = new JLabel("��",JLabel.CENTER);
        ro_lab.setBounds(120, 20, 100, 30);
        ro_lab.setLocale(new Locale("ja","JP"));
        ro_lab.setFont(new java.awt.Font("dialog", 0, 18));
        ro_lab.setBorder(new Flush3DBorder());
        ro_lab.setForeground(java.awt.Color.black);
        getContentPane().add(ro_lab);

        copy_lab = new JLabel("��",JLabel.CENTER);
        copy_lab.setBounds(45, 70, 150, 30);
        copy_lab.setLocale(new Locale("ja","JP"));
        copy_lab.setFont(new java.awt.Font("dialog", 0, 18));
        copy_lab.setBorder(new Flush3DBorder());
        copy_lab.setForeground(java.awt.Color.black);
        getContentPane().add(copy_lab);

        lab = new JLabel("�R�s�[��",JLabel.CENTER);
        lab.setBounds(20, 120, 100, 30);
        lab.setLocale(new Locale("ja","JP"));
        lab.setFont(new java.awt.Font("dialog", 0, 18));
        lab.setBorder(new Flush3DBorder());
        lab.setForeground(java.awt.Color.black);
        getContentPane().add(lab);

        ro_name = new RoText();
        ro_name.setBounds(120, 120, 100, 30);
        getContentPane().add(ro_name);

        lab = new JLabel("�ݒ��",JLabel.CENTER);
        lab.setBounds(20, 400, 100, 24);
        lab.setLocale(new Locale("ja","JP"));
        lab.setFont(new java.awt.Font("dialog", 0, 16));
        lab.setBorder(new Flush3DBorder());
        lab.setForeground(java.awt.Color.black);
        getContentPane().add(lab);

        op_name = new TText();
        op_name.setBounds(120, 400, 140, 24);
        getContentPane().add(op_name);

        send_button = new JButton("��  �s");
        send_button.setBounds(260, 400, 100, 24);
        send_button.setLocale(new Locale("ja","JP"));
        send_button.setFont(new java.awt.Font("dialog", 0, 18));
        send_button.setBorder(new Flush3DBorder());
        send_button.setForeground(java.awt.Color.black);
        send_button.addActionListener(new SendButton());
        getContentPane().add(send_button);

        cancel_button = new JButton("�I  ��");
        cancel_button.setBounds(380, 400, 100, 24);
        cancel_button.setLocale(new Locale("ja","JP"));
        cancel_button.setFont(new java.awt.Font("dialog", 0, 18));
        cancel_button.setBorder(new Flush3DBorder());
        cancel_button.setForeground(java.awt.Color.black);
        cancel_button.addActionListener(new CancelButton());
        getContentPane().add(cancel_button);

        top = new DefaultMutableTreeNode("���ƒ萔");

        for(int i = 0 ; ; i++){ // for 1
            CZSystemOpTbLag largename = CZSystem.getOpTbLag(i);
            if(null == largename) break;
            LargeNode large = new LargeNode(largename.k_name.trim(),largename);
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
        treepanel.setBounds(240, 20, 240, 354);
        treepanel.setBorder(new Flush3DBorder());
        treepanel.setForeground(java.awt.Color.black);
        getContentPane().add(treepanel);

    }


    //
    //
    //
    private void selectAllData(){
//@@        CZSystem.log("CZOperationTableCp","selectAllData()");

        current_lag = -1;
        current_mid = -1;
        copy_mode = ALL_COPY;
    }


    //
    //
    //
    private void selectLagData(int lag){
//@@        CZSystem.log("CZOperationTableCp","selectLagData(" + lag + ")");

        current_lag = lag;
        current_mid = -1;
        copy_mode = LARGE_COPY;
    }

    //
    //
    //
    private void selectMidData(int lag , int mid ){
//@@        CZSystem.log("CZOperationTableCp","selectMidData(" + lag + ")(" + mid + ")");

        current_lag = lag;
        current_mid = mid;
        copy_mode = MIDDLE_COPY;
    }


    //
    //
    //
    private boolean sendChk(){

        if(1 > ro_name.getText().length()){
//@@            CZSystem.log("CZOperationTableCp","setSendStatus() Table Ro Name Error !!");
            Object msg[] = {"���ƒ萔�R�s�[",
                                "�F������͂��Ă������I�I",
                            ""};
            errorMsg(msg);
            return false;
        }

        if(1 > op_name.getText().length()){
//@@            CZSystem.log("CZOperationTableCp","setSendStatus() Table Op Name Error !!");
            Object msg[] = {"���ƒ萔�R�s�[",
                            "�ݒ�҂���͂��Ă������I�I",
                            ""};
            errorMsg(msg);
            return false;
        }

        if(NO_COPY == copy_mode){
//@@            CZSystem.log("CZOperationTableCp","setSendStatus() Table Copy Mode Error !!");
            Object msg[] = {"���ƒ萔�R�s�[",
                            "�I�����������Ă��������I�I",
                            ""};
            errorMsg(msg);
            return false;
        }
        return true;
    }


    //
    //
    //
    public boolean setDefault(){

//@@        CZSystem.log("CZOperationTableCp","setDefault()");
        current_lag = -1;
        current_mid = -1;
        copy_mode   = NO_COPY;

		String s = CZSystem.RoKetaChg(CZSystem.getRoName());	// 20050725 �F�F�\�������ύX
        ro_lab.setText(s);
//        ro_lab.setText(CZSystem.getRoName());

        if(CZSystemDefine.ADMIN_RUN != CZSystem.getRunLevel()){
            send_button.setEnabled(false);
        }

        op_name.setText("");
        return true;
    }

    //
    // ���b�Z�[�W�̕\��
    //
    private boolean errorMsg(Object msg[]){
        JOptionPane.showMessageDialog(null,msg,
            "���ƒ萔���̓G���[",
            JOptionPane.ERROR_MESSAGE);
        return true;
    }


    /*
    *
    *
    *
    */
    class SendButton implements ActionListener {
        public void actionPerformed(ActionEvent ev){

            if(!sendChk()) return;
            boolean ret = false;
            String  copy_st = "���I��";

			String ro = new String();
				if( 0 != CZSystemDefine.DISP_KETA_FLG){
					StringBuffer a = new StringBuffer();
					a.append(ro_name.getText());
					a.insert(0,"K");
					String s = a.toString();
					ro = s;
				} else {
					ro = ro_name.getText();
				}
//            String ro = ro_name.getText();
            String op = op_name.getText();

            switch(copy_mode){
                case ALL_COPY :     // �F�ԃR�s�[
                    ret = CZSystem.CZWorkingCopyRo(op,ro);
                    copy_st = "�F�ԃR�s�[";
                    break;
                            
                case LARGE_COPY :   // �區�ڃR�s�[
                    ret = CZSystem.CZWorkingCopyNo1(op,ro,current_lag);
                    copy_st = "�區�ڃR�s�[";
                    break;

                case MIDDLE_COPY :  // �����ڃR�s�[
                    ret = CZSystem.CZWorkingCopyNo2(op,ro,current_lag,current_mid);
                    copy_st = "�����ڃR�s�[";
                    break;
            }

            if(!ret){
//@@                CZSystem.log("CZOperationTableCp","setSendStatus() Table Data Lag Error !!");
                Object msg1[] = {"���ƒ萔�R�s�[",
                                "  [" + copy_st + "] �G���[�I�I"};

                CZSystemSysMsg msg2 = new CZSystemSysMsg();
                msg2.no = -1;
                msg2.message = CZSystem.getDateTime() + "  ���ƒ萔�R�s�[[" + copy_st + "]�G���[";
                CZSystem.sysMessage(msg2);

                errorMsg(msg1);
            }
            else{
                CZSystemSysMsg msg3 = new CZSystemSysMsg();
                msg3.no = 0;
                msg3.message = CZSystem.getDateTime() + "  ���ƒ萔�R�s�[[" + copy_st + "]";
                CZSystem.sysMessage(msg3);
            }
        }
    }


    /*
    *
    *
    *
    */
    class CancelButton implements ActionListener {
        public void actionPerformed(ActionEvent ev){
            setDefault();
            setVisible(false);
        }
    }


    /*
    *
    *
    *
    */
    class LargeNode extends DefaultMutableTreeNode {
        private CZSystemOpTbLag data = null;

        LargeNode(String name,CZSystemOpTbLag dat){
            super(name);
            data = dat;
        }

        public CZSystemOpTbLag getData(){
            return data;
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

            copy_mode = NO_COPY;
            copy_lab.setText("��");

            if(node.isRoot()){
//@@                CZSystem.log("CZOperationTableCp","���ƒ萔�S�R�s�[");
                copy_lab.setText("�S���ڃR�s�[");

                selectAllData();
                return;
            }

            if(node.isLeaf()){
                Node n = (Node)node;
                CZSystemOpTbMid dat = n.getData();

//@@                CZSystem.log("CZOperationTableCp",
//@@                    "���ƒ萔���R�s�[ [" + dat.k_no1 + "][" + dat.k_no2 + "][" + dat.k_name + "]");
                copy_lab.setText("�����ڃR�s�[");
                selectMidData(dat.k_no1,dat.k_no2);
                return;
            }

            if(node.getRoot() == node.getParent()){
                LargeNode n = (LargeNode)node;
                CZSystemOpTbLag dat = n.getData();
                    
//@@                CZSystem.log("CZOperationTableCp","���ƒ萔��R�s�[ [" + dat.k_no + "][" + dat.k_name + "]");
                copy_lab.setText("�區�ڃR�s�[");
                selectLagData(dat.k_no);
                return;
            }
            CZSystem.log("CZOperationTableCp","���ƒ萔�R�s�[ Node Error]");
            setDefault();
        }
    }


    /*
    *
    *       �ݒ�҂���͂���TextField
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


    /**
    *
    *       �F����͂���TextField
    *
    */
    public class RoText extends JTextField {

        RoText(){
            super();
            setFont(new java.awt.Font("dialog", 0, 18));
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

            String validValues = "0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ";

            //
            //
            public void insertString( int offset, String str, AttributeSet a )
                                    throws BadLocationException {
                if(5 < getLength()) return;
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
}
