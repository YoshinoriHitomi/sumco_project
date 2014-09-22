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
 *   MO�Ǘ�Window
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

        setTitle("�l�n�Ǘ�");
        setSize(460,140);
        setResizable(false);
        setModal(true);

        getContentPane().setLayout(null);
        // ����n�Q�Ƌ@�\    @20131021
        if(CZSystemDefine.REFERENCE_RUN == CZSystem.getRunLevel()){
            getContentPane().setBackground(CZSystemDefine.DEFAULT_REFERENCE_BACKGROUND_COL);
        }else{
            getContentPane().setBackground(CZSystemDefine.DEFAULT_BACKGROUND_COL);
        }

        mo1_button = new JButton("�l�n�P");
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

        mo2_button = new JButton("�l�n�Q");
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
    // �m�F���b�Z�[�W�̕\��
    //
    private boolean confirmDia(Object msg[]){

        int ans = JOptionPane.showConfirmDialog(null,msg,
              "�l�n�Ǘ�",
              JOptionPane.OK_CANCEL_OPTION,
              JOptionPane.WARNING_MESSAGE);

        if(0 == ans) return true;
        return false;
    }


    //
    // �G���[���b�Z�[�W�̕\��
    //
    private boolean errorMsg(Object msg[]){
        JOptionPane.showMessageDialog(null,msg,
        "�l�n�Ǘ��G���[",
        JOptionPane.ERROR_MESSAGE);
        return true;
    }


    //
    // ���b�Z�[�W�̕\��
    //
    private boolean showMsg(Object msg[]){
        JOptionPane.showMessageDialog(null,msg,
        "�l�n�Ǘ�",
        JOptionPane.INFORMATION_MESSAGE);
        return true;
    }

    /*
    *
    *   �l�n�̏�Ԏ擾
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
                Object msg[] = {"�l�n�̏�Ԏ擾�Ɏ��s���܂����I�I"};
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
    *   �l�n��Mount
    *
    */
    class MOMountButton implements ActionListener {
        private int dev_no = 0;

        MOMountButton(int no){
            dev_no = no;
        }

        public void actionPerformed(ActionEvent ev){
            if(1 > dev_no) return;

            String m = new String("  1) MO [" + dev_no + "] �ł����H");
            Object msg1[] = {"MO �� Mount ���J�n���܂��B���L�̍��ڂ��m�F���Ă��������I�I",
                        m,
                        "  2) MO������Mount����Ă܂��񂩁H",
                        "  3) MO���}������Ă܂����H"};

            if(!confirmDia(msg1)) return;

            //Send
            boolean ret = CZSystem.CZMoMount(dev_no);
            if(!ret){
                Object msg2[] = {"�l�n��Mount�Ɏ��s���܂����I�I"};
                errorMsg(msg2);
                return;
            }

            Object msg3[] = {"�l�n��Mount�ɐ������܂����I�I"};
            showMsg(msg3);
        }
    }


    /*
    *
    *
    *   �l�n��UMount
    *
    */
    class MOUmountButton implements ActionListener {
        private int dev_no = 0;

        MOUmountButton(int no){
            dev_no = no;
        }

        public void actionPerformed(ActionEvent ev){
            if(1 > dev_no) return;

            String m = new String("  1) MO [" + dev_no + "] �ł����H");
            Object msg1[] = {"MO �� UnMount ���J�n���܂��B���L�̍��ڂ��m�F���Ă��������I�I",
                        m,
                        "  2) MO������UnMount����Ă܂��񂩁H",
                        "  3) MO�ɏ������ݒ��ł͂���܂��񂩁H"};

            if(!confirmDia(msg1)) return;

            //Send
            boolean ret = CZSystem.CZMoUmount(dev_no);
            if(!ret){
                Object msg2[] = {"�l�n��UnMount�Ɏ��s���܂����I�I"};
                errorMsg(msg2);
                return;
            }

            Object msg3[] = {"�l�n��UnMount�ɐ������܂����I�I"};
            showMsg(msg3);
        }
    }


    /*
    *
    *
    *   �l�n��Format
    *
    */
    class MOFormatButton implements ActionListener {
        private int dev_no = 0;

        MOFormatButton(int no){
            dev_no = no;
        }

        public void actionPerformed(ActionEvent ev){
            if(1 > dev_no) return;

            String m = new String("  1) MO [" + dev_no + "] �ł����H");
            Object msg1[] = {"MO �� Format ���J�n���܂��B���L�̍��ڂ��m�F���Ă��������I�I",
                        m,
                        "  2) MO���̃f�[�^�͕s�v�ł����H",
                        "  3) MO��UnMount����Ă��܂����H",
                        "  4) MO���}������Ă܂����H"};

            if(!confirmDia(msg1)) return;

            Object msg2[] = {"�Ō�̊m�F�ł��B�f�[�^�͑S�Ď����܂�",
                        "�{���ɂ�낵���ł����H"};

            if(!confirmDia(msg2)) return;

            //Send
            boolean ret = CZSystem.CZMoFormat(dev_no);
            if(!ret){
                Object msg3[] = {"�l�n��Format�Ɏ��s���܂����I�I"};
                errorMsg(msg3);
                return;
            }

            Object msg4[] = {"�l�n��Format�ɐ������܂����I�I"};
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

            setTitle("�l�n�f�[�^�c�a�W�J");
            setSize(870,275);
            setResizable(false);
            setModal(true);

            getContentPane().setLayout(null);
            // ����n�Q�Ƌ@�\    @20131021
            if(CZSystemDefine.REFERENCE_RUN == CZSystem.getRunLevel()){
                getContentPane().setBackground(CZSystemDefine.DEFAULT_REFERENCE_BACKGROUND_COL);
            }else{
                getContentPane().setBackground(CZSystemDefine.DEFAULT_BACKGROUND_COL);
            }

            dir_scpanel = new JScrollPane();
            dir_scpanel.setBounds(20, 20, 830, 187);
            getContentPane().add(dir_scpanel);

            send_button = new JButton("��  �s");
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
                    Object msg1[] = {"�Ώۂ�I�����Ă��������I�I"};
                    errorMsg(msg1);
                    return;
                }

                //Send
                boolean ret = CZSystem.CZMoExtract(mount_point,dir_list[row].getRoName(),dir_list[row].getBatch());
                if(!ret){
                    Object msg2[] = {"�l�n�̂c�a�W�J�Ɏ��s���܂����I�I"};
                    errorMsg(msg2);
                    return;
                }

                Object msg3[] = {"�l�n�̂c�a�W�J�ɐ������܂����I�I"};
                showMsg(msg3);
            }
        }


        /*
        *
        *   �l�n�ۑ��ꗗ
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

                    // �F
                    colum = cmdl.getColumn(1);
                    colum.setMaxWidth(60);
                    colum.setMinWidth(60);
                    colum.setWidth(60);

                    // BtNo
                    colum = cmdl.getColumn(2);
                    colum.setMaxWidth(130);
                    colum.setMinWidth(130);
                    colum.setWidth(130);


                    // �t�@�C��
                    colum = cmdl.getColumn(3);
                    colum.setMaxWidth(200);
                    colum.setMinWidth(200);
                    colum.setWidth(200);

                    // �J�n����
                    colum = cmdl.getColumn(4);
                    colum.setMaxWidth(170);
                    colum.setMinWidth(170);
                    colum.setWidth(170);

                    // �I������
                    colum = cmdl.getColumn(5);
                    colum.setMaxWidth(170);
                    colum.setMinWidth(170);
                    colum.setWidth(170);

                    // �Ԉ���
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
        *       �l�n�ۑ��ꗗ�F���f��
        *
        */

        public class MOTblMdl extends AbstractTableModel {

            final   int TBL_COL = 7;
            private int TBL_ROW = 0;

            final String[] names = {" # " ,
                                "�F",
                                "BtNo",
                                "�t�@�C����",
                                "�J�n����",
                                "�I������",
                                "�Ԉ���"};

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
                        case 0: data[i][6] = new String("�L��");
                            break;
                        case 1: data[i][6] = new String("����");
                            break;
                        default:data[i][6] = new String("�s��");
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

