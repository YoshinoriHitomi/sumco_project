package cz;

import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.util.Locale;

import javax.swing.JButton;
import javax.swing.JDialog;
import javax.swing.JOptionPane;
import javax.swing.JScrollPane;
import javax.swing.JTable;
import javax.swing.ListSelectionModel;
import javax.swing.event.ListSelectionEvent;
import javax.swing.plaf.metal.MetalBorders.Flush3DBorder;
import javax.swing.table.AbstractTableModel;
import javax.swing.table.DefaultTableColumnModel;
import javax.swing.table.JTableHeader;
import javax.swing.table.TableColumn;

import czclass.CZMoList;


/**
 *   MO管理Window
 * @author  (KPK Co.,Ltd.)
 * @version 1.0 (2003/04/01)
 *
 */

public class CZMOControl extends JDialog {

    private JButton     mo1_button        = null;
    private JButton     mo1_mount_button  = null;
    private JButton     mo1_umount_button = null;
    private JButton     mo1_format_button = null;

    private JButton     mo2_button        = null;
    private JButton     mo2_mount_button  = null;
    private JButton     mo2_umount_button = null;
    private JButton     mo2_format_button = null;

    private MODirs      mo_dirs           = null; 

    //
    //
    //
    CZMOControl(){
        super();

        setTitle("ＭＯ管理");
        setSize(460,140);
        setResizable(false);
        setModal(true);

        getContentPane().setLayout(null);
        // 他基地参照機能    @20131021
        if(CZSystemDefine.REFERENCE_RUN == CZSystem.getRunLevel()){
            getContentPane().setBackground(CZSystemDefine.DEFAULT_REFERENCE_BACKGROUND_COL);
        }else{
            getContentPane().setBackground(CZSystemDefine.DEFAULT_BACKGROUND_COL);
        }

        mo1_button = new JButton("ＭＯ１");
        mo1_button.setBounds(20, 20, 100, 24);
        mo1_button.setLocale(new Locale("ja","JP"));
        mo1_button.setFont(new java.awt.Font("dialog", 0, 18));
        mo1_button.setBorder(new Flush3DBorder());
        mo1_button.setForeground(java.awt.Color.black);
        mo1_button.addActionListener(new MOButton(1));
        getContentPane().add(mo1_button);

        mo1_mount_button = new JButton("Mount");
        mo1_mount_button.setBounds(120, 20, 100, 24);
        mo1_mount_button.setLocale(new Locale("ja","JP"));
        mo1_mount_button.setFont(new java.awt.Font("dialog", 0, 18));
        mo1_mount_button.setBorder(new Flush3DBorder());
        mo1_mount_button.setForeground(java.awt.Color.black);
        mo1_mount_button.addActionListener(new MOMountButton(1));
        getContentPane().add(mo1_mount_button);

        mo1_umount_button = new JButton("Umount");
        mo1_umount_button.setBounds(220, 20, 100, 24);
        mo1_umount_button.setLocale(new Locale("ja","JP"));
        mo1_umount_button.setFont(new java.awt.Font("dialog", 0, 18));
        mo1_umount_button.setBorder(new Flush3DBorder());
        mo1_umount_button.setForeground(java.awt.Color.black);
        mo1_umount_button.addActionListener(new MOUmountButton(1));
        getContentPane().add(mo1_umount_button);

        mo1_format_button = new JButton("Format");
        mo1_format_button.setBounds(340, 20, 100, 24);
        mo1_format_button.setLocale(new Locale("ja","JP"));
        mo1_format_button.setFont(new java.awt.Font("dialog", 0, 18));
        mo1_format_button.setBorder(new Flush3DBorder());
        mo1_format_button.setForeground(java.awt.Color.black);
        mo1_format_button.addActionListener(new MOFormatButton(1));
        getContentPane().add(mo1_format_button);

        //////////////////////////////////////////////////////////////////////

        mo2_button = new JButton("ＭＯ２");
        mo2_button.setBounds(20, 70, 100, 24);
        mo2_button.setLocale(new Locale("ja","JP"));
        mo2_button.setFont(new java.awt.Font("dialog", 0, 18));
        mo2_button.setBorder(new Flush3DBorder());
        mo2_button.setForeground(java.awt.Color.black);
        mo2_button.addActionListener(new MOButton(2));
        getContentPane().add(mo2_button);

        mo2_mount_button = new JButton("Mount");
        mo2_mount_button.setBounds(120, 70, 100, 24);
        mo2_mount_button.setLocale(new Locale("ja","JP"));
        mo2_mount_button.setFont(new java.awt.Font("dialog", 0, 18));
        mo2_mount_button.setBorder(new Flush3DBorder());
        mo2_mount_button.setForeground(java.awt.Color.black);
        mo2_mount_button.addActionListener(new MOMountButton(2));
        getContentPane().add(mo2_mount_button);

        mo2_umount_button = new JButton("Umount");
        mo2_umount_button.setBounds(220, 70, 100, 24);
        mo2_umount_button.setLocale(new Locale("ja","JP"));
        mo2_umount_button.setFont(new java.awt.Font("dialog", 0, 18));
        mo2_umount_button.setBorder(new Flush3DBorder());
        mo2_umount_button.setForeground(java.awt.Color.black);
        mo2_umount_button.addActionListener(new MOUmountButton(2));
        getContentPane().add(mo2_umount_button);

        mo2_format_button = new JButton("Format");
        mo2_format_button.setBounds(340, 70, 100, 24);
        mo2_format_button.setLocale(new Locale("ja","JP"));
        mo2_format_button.setFont(new java.awt.Font("dialog", 0, 18));
        mo2_format_button.setBorder(new Flush3DBorder());
        mo2_format_button.setForeground(java.awt.Color.black);
        mo2_format_button.addActionListener(new MOFormatButton(2));
        getContentPane().add(mo2_format_button);

        mo_dirs = new MODirs();
        mo_dirs.setVisible(false);
    }


    //
    //
    //
    public boolean setDefault(){

        mo1_button.setEnabled(false);
        mo1_mount_button.setEnabled(false);
        mo1_umount_button.setEnabled(false);
        mo1_format_button.setEnabled(false);

        mo2_button.setEnabled(false);
        mo2_mount_button.setEnabled(false);
        mo2_umount_button.setEnabled(false);
        mo2_format_button.setEnabled(false);

        if(CZSystemDefine.ADMIN_RUN == CZSystem.getRunLevel()){
            mo1_button.setEnabled(true);
            mo1_mount_button.setEnabled(true);
            mo1_umount_button.setEnabled(true);
            mo1_format_button.setEnabled(true);

            mo2_button.setEnabled(true);
            mo2_mount_button.setEnabled(true);
            mo2_umount_button.setEnabled(true);
            mo2_format_button.setEnabled(true);
        }
        return true;
    }



    //
    // 確認メッセージの表示
    //
    private boolean confirmDia(Object msg[]){

        int ans = JOptionPane.showConfirmDialog(null,msg,
              "ＭＯ管理",
              JOptionPane.OK_CANCEL_OPTION,
              JOptionPane.WARNING_MESSAGE);

        if(0 == ans) return true;
        return false;
    }


    //
    // エラーメッセージの表示
    //
    private boolean errorMsg(Object msg[]){
        JOptionPane.showMessageDialog(null,msg,
        "ＭＯ管理エラー",
        JOptionPane.ERROR_MESSAGE);
        return true;
    }


    //
    // メッセージの表示
    //
    private boolean showMsg(Object msg[]){
        JOptionPane.showMessageDialog(null,msg,
        "ＭＯ管理",
        JOptionPane.INFORMATION_MESSAGE);
        return true;
    }

    /*
    *
    *   ＭＯの状態取得
    *
    */
    class MOButton implements ActionListener {
        private int dev_no = 0;

        MOButton(int no){
            dev_no = no;
        }

        public void actionPerformed(ActionEvent ev){

//@@            CZSystem.log("CZMOControl MOButton","DevNo[" + dev_no + "]");

            if(1 > dev_no) return;

            //Send
            CZMoList[] list = CZSystem.CZMoGetlist(dev_no);


            if(null == list){
                Object msg[] = {"ＭＯの状態取得に失敗しました！！"};
                errorMsg(msg);
                return;
            }

            mo_dirs.setDefault(dev_no,list);

            mo_dirs.setVisible(true);

        }
    }


    /*
    *
    *
    *   ＭＯのMount
    *
    */
    class MOMountButton implements ActionListener {
        private int dev_no = 0;

        MOMountButton(int no){
            dev_no = no;
        }

        public void actionPerformed(ActionEvent ev){
            if(1 > dev_no) return;

            String m = new String("  1) MO [" + dev_no + "] ですか？");
            Object msg1[] = {"MO の Mount を開始します。下記の項目を確認してください！！",
                        m,
                        "  2) MOが既にMountされてませんか？",
                        "  3) MOが挿入されてますか？"};

            if(!confirmDia(msg1)) return;

            //Send
            boolean ret = CZSystem.CZMoMount(dev_no);
            if(!ret){
                Object msg2[] = {"ＭＯのMountに失敗しました！！"};
                errorMsg(msg2);
                return;
            }

            Object msg3[] = {"ＭＯのMountに成功しました！！"};
            showMsg(msg3);
        }
    }


    /*
    *
    *
    *   ＭＯのUMount
    *
    */
    class MOUmountButton implements ActionListener {
        private int dev_no = 0;

        MOUmountButton(int no){
            dev_no = no;
        }

        public void actionPerformed(ActionEvent ev){
            if(1 > dev_no) return;

            String m = new String("  1) MO [" + dev_no + "] ですか？");
            Object msg1[] = {"MO の UnMount を開始します。下記の項目を確認してください！！",
                        m,
                        "  2) MOが既にUnMountされてませんか？",
                        "  3) MOに書き込み中ではありませんか？"};

            if(!confirmDia(msg1)) return;

            //Send
            boolean ret = CZSystem.CZMoUmount(dev_no);
            if(!ret){
                Object msg2[] = {"ＭＯのUnMountに失敗しました！！"};
                errorMsg(msg2);
                return;
            }

            Object msg3[] = {"ＭＯのUnMountに成功しました！！"};
            showMsg(msg3);
        }
    }


    /*
    *
    *
    *   ＭＯのFormat
    *
    */
    class MOFormatButton implements ActionListener {
        private int dev_no = 0;

        MOFormatButton(int no){
            dev_no = no;
        }

        public void actionPerformed(ActionEvent ev){
            if(1 > dev_no) return;

            String m = new String("  1) MO [" + dev_no + "] ですか？");
            Object msg1[] = {"MO の Format を開始します。下記の項目を確認してください！！",
                        m,
                        "  2) MO中のデータは不要ですか？",
                        "  3) MOはUnMountされていますか？",
                        "  4) MOが挿入されてますか？"};

            if(!confirmDia(msg1)) return;

            Object msg2[] = {"最後の確認です。データは全て失われます",
                        "本当によろしいですか？"};

            if(!confirmDia(msg2)) return;

            //Send
            boolean ret = CZSystem.CZMoFormat(dev_no);
            if(!ret){
                Object msg3[] = {"ＭＯのFormatに失敗しました！！"};
                errorMsg(msg3);
                return;
            }

            Object msg4[] = {"ＭＯのFormatに成功しました！！"};
            showMsg(msg4);
        }
    }


    /*
    *
    */
    class MODirs extends JDialog {
        private JButton     send_button = null;

        private JScrollPane dir_scpanel = null;
        private MOTable     dir_table   = null;

        private int     mount_point = -1;
        private CZMoList[]  dir_list    = null;

        MODirs(){
            super();

            setTitle("ＭＯデータＤＢ展開");
            setSize(870,275);
            setResizable(false);
            setModal(true);

            getContentPane().setLayout(null);
            // 他基地参照機能    @20131021
            if(CZSystemDefine.REFERENCE_RUN == CZSystem.getRunLevel()){
                getContentPane().setBackground(CZSystemDefine.DEFAULT_REFERENCE_BACKGROUND_COL);
            }else{
                getContentPane().setBackground(CZSystemDefine.DEFAULT_BACKGROUND_COL);
            }

            dir_scpanel = new JScrollPane();
            dir_scpanel.setBounds(20, 20, 830, 187);
            getContentPane().add(dir_scpanel);

            send_button = new JButton("実  行");
            send_button.setBounds(20, 220, 100, 24);
            send_button.setLocale(new Locale("ja","JP"));
            send_button.setFont(new java.awt.Font("dialog", 0, 18));
            send_button.setBorder(new Flush3DBorder());
            send_button.setForeground(java.awt.Color.black);
            send_button.addActionListener(new SendButton());
            getContentPane().add(send_button);

        }


        //
        //
        //
        public boolean setDefault(int no,CZMoList[] l){

            mount_point = no;
            dir_list = l;

            for (int i=0; i<dir_list.length; i++) {
                System.out.println ("MODEV:" + mount_point);
                System.out.println ("   ro:" + dir_list[i].getRoName());
                System.out.println ("batch:" + dir_list[i].getBatch());
                System.out.println (" file:" + dir_list[i].getTgzFile());
                System.out.println ("stime:" + dir_list[i].getStime());
                System.out.println ("etime:" + dir_list[i].getEtime());
                System.out.println ("  flg:" + dir_list[i].getMflg());
                System.out.println (" ");
            }

            dir_table = new MOTable(dir_list);
            JTableHeader tabHead = dir_table.getTableHeader();
            tabHead.setReorderingAllowed(false);

            dir_scpanel.setViewportView(dir_table);

            return true;
        }

        /*
        *
        *
        *
        */
        class SendButton implements ActionListener {
            public void actionPerformed(ActionEvent ev){

                int row = dir_table.getSelectedRow();

                System.out.println ("SendButton Row:" + row);

                if(0 > row){
                    Object msg1[] = {"対象を選択してください！！"};
                    errorMsg(msg1);
                    return;
                }

                //Send
                boolean ret = CZSystem.CZMoExtract(mount_point,dir_list[row].getRoName(),dir_list[row].getBatch());
                if(!ret){
                    Object msg2[] = {"ＭＯのＤＢ展開に失敗しました！！"};
                    errorMsg(msg2);
                    return;
                }

                Object msg3[] = {"ＭＯのＤＢ展開に成功しました！！"};
                showMsg(msg3);
            }
        }


        /*
        *
        *   ＭＯ保存一覧
        *
        */
        class MOTable extends JTable {

            private MOTblMdl model = null;

            MOTable(CZMoList[] l){
                super();

                try{
                    setName("MOTable");
                    setBounds(0, 0, 200, 200);
                    setAutoCreateColumnsFromModel(true);
                    setSelectionMode(ListSelectionModel.SINGLE_SELECTION);

                    setLocale(new Locale("ja","JP"));
                    setFont(new java.awt.Font("dialog", 0, 12));
                    setRowHeight(17);

                    if(null == l) return;
                    if(1 > l.length) return;


                    model = new MOTblMdl(l);
                    setModel(model);


                    DefaultTableColumnModel cmdl = (DefaultTableColumnModel)getColumnModel();
                    TableColumn  colum = null;


                    // No
                    colum = cmdl.getColumn(0);
                    colum.setMaxWidth(40);
                    colum.setMinWidth(40);
                    colum.setWidth(40);

                    // 炉
                    colum = cmdl.getColumn(1);
                    colum.setMaxWidth(60);
                    colum.setMinWidth(60);
                    colum.setWidth(60);

                    // BtNo
                    colum = cmdl.getColumn(2);
                    colum.setMaxWidth(130);
                    colum.setMinWidth(130);
                    colum.setWidth(130);


                    // ファイル
                    colum = cmdl.getColumn(3);
                    colum.setMaxWidth(200);
                    colum.setMinWidth(200);
                    colum.setWidth(200);

                    // 開始日時
                    colum = cmdl.getColumn(4);
                    colum.setMaxWidth(170);
                    colum.setMinWidth(170);
                    colum.setWidth(170);

                    // 終了日時
                    colum = cmdl.getColumn(5);
                    colum.setMaxWidth(170);
                    colum.setMinWidth(170);
                    colum.setWidth(170);

                    // 間引き
                    colum = cmdl.getColumn(6);
                    colum.setMaxWidth(40);
                    colum.setMinWidth(40);
                    colum.setWidth(40);

                }
                catch (Throwable e) {
                    CZSystem.handleException(e);
                }
                return;
            }

            //
            //
            //
            public void valueChanged(ListSelectionEvent e){
                super.valueChanged(e);

                System.out.println("MODirs MOTable valueChanged [" + getSelectedRow() + "][" + getSelectedColumn() + "]");
            }
        } // MOTable

        /*
        *
        *       ＭＯ保存一覧：モデル
        *
        */

        public class MOTblMdl extends AbstractTableModel {

            final   int TBL_COL = 7;
            private int TBL_ROW = 0;

            final String[] names = {" # " ,
                                "炉",
                                "BtNo",
                                "ファイル名",
                                "開始日時",
                                "終了日時",
                                "間引き"};

            private Object  data[][];

            MOTblMdl(CZMoList[] l){
                super();

                TBL_ROW = l.length;

                data = new Object[TBL_ROW][TBL_COL];
                    
                for(int i = 0 ; i < TBL_ROW ; i++){
                    data[i][0] = new Integer(i+1);
                    data[i][1] = l[i].getRoName();
                    data[i][2] = l[i].getBatch();
                    data[i][3] = l[i].getTgzFile();
                    data[i][4] = l[i].getStime();
                    data[i][5] = l[i].getEtime();

                    switch(l[i].getMflg()){
                        case 0: data[i][6] = new String("有り");
                            break;
                        case 1: data[i][6] = new String("無し");
                            break;
                        default:data[i][6] = new String("不明");
                            break;
                    }
                } // for end
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

        } // GrTblMdl
    } // MODirs
} // CZMOControl

