package cz;

import java.awt.Component;
import java.awt.Cursor;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.PrintWriter;
import java.io.Serializable;
import java.util.Locale;
import java.util.Properties;
import java.util.Vector;

import javax.swing.JButton;
import javax.swing.JComboBox;
import javax.swing.JDialog;
import javax.swing.JLabel;
import javax.swing.JOptionPane;
import javax.swing.JPanel;
import javax.swing.JScrollPane;
import javax.swing.JTable;
import javax.swing.JTextField;
import javax.swing.JViewport;
import javax.swing.JFileChooser;
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

/*******************************************************************************
 *  �o�u�f�[�^�ۑ�
 * @author  (KPK Co.,Ltd.)
 * @version 1.0 (2003/04/01)
 * T6�ǉ��ɔ����C�� @@
********************************************************************************/
public class CZPVDataSave extends JDialog {
	int		iIdo = 44+4;

    private static CZPVDataSave pvdatasave      = null;

    private JScrollPane bt_scpanel              = null;
    private JScrollPane bt_start_scpanel        = null;
    private JScrollPane bt_condition_scpanel    = null;

    private JPanel      pv_panel        = null;
    private PVTable     pv_table        = null;

    private JButton     save_button     = null;
    private JButton     mabiki_button     = null;		/* 2003.10.20 y.k */


    private String prop_Unit     = null;
    private String prop_Start    = null;
    private String prop_End      = null;
    private String prop_Interval = null;

    private String prop_PVItemNo[];

    private JButton     set_save_button     = null;		/* 2008.09.18 */
    private JButton     set_read_button     = null;		/* 2008.09.18 */

    private Vector      start_list      = null;

    private JComboBox   unit_cbox       = null;

    private NumeralText start_text      = null;
    private NumeralText end_text        = null;
    private NumeralText interval_text   = null;

    private JLabel      start_lab       = null;
    private JLabel      end_lab         = null;
    private JLabel      interval_lab    = null;

    private FileText    file_text       = null;

	private BtTable		btTable = null;

    private String      save_dir        = null;

    private File        file_ = new File(CZSystem.FILE_SRC_PATH);

    //
    // �R���X�g���N�^
    //
    CZPVDataSave(){
        super();

        setTitle("�o�u�f�[�^�ۑ�");
        setSize(820+120,705+iIdo);
        setResizable(false);
        setModal(true);

        getContentPane().setLayout(null);
        // ����n�Q�Ƌ@�\    @20131021
        if(CZSystemDefine.REFERENCE_RUN == CZSystem.getRunLevel()){
            getContentPane().setBackground(CZSystemDefine.DEFAULT_REFERENCE_BACKGROUND_COL);
        }else{
            getContentPane().setBackground(CZSystemDefine.DEFAULT_BACKGROUND_COL);
        }

        try{
            Properties prop =  new Properties();
            FileInputStream pros = new FileInputStream(CZSystemDefine.PROPERTY_FILE);
            prop.load(pros);

            save_dir = prop.getProperty("PV_SAVE_FILE_DIR");
            if(null == save_dir) CZSystem.exit(-1,"CZPVDataSave NO Propertie File null");
            if(1 > save_dir.length()) CZSystem.exit(-1,"CZPVDataSave NO Propertie File name");
        }
        catch( Exception e){
            CZSystem.exit(-1,"CZPVDataSave NO Propertie File");
        }

        try{
            // ----- Property_File��� �I��PV���ځA�\���ݒ���擾����B --------
            Properties prop =  new Properties();
            FileInputStream pros = new FileInputStream("PVITEM.TXT");
            prop.load(pros);

            // PV���ڂ̐ݒ�
            prop_PVItemNo  = new String[128];
            for(int i=0; i < 128 ; i++){
                try {
                    prop_PVItemNo[i]   = prop.getProperty("PVDATA" + (i+1));
                } catch (Exception e) {
                    prop_PVItemNo[i]   = new String("");
                }
            }
            // �\���ݒ�
            prop_Unit     = prop.getProperty("SELUNIT");
            prop_Start    = prop.getProperty("START");
            prop_End      = prop.getProperty("END");
            prop_Interval = prop.getProperty("INTERVAL");

        } catch( Exception e ) {
                                        //�v���p�e�B�擾�ŃG���[�̎��́A�I������B
            CZSystem.exit(-1,"CZPVDataSave NO Propertie File");
        }

		/* ����i�̃p�l�� */
        bt_scpanel = new JScrollPane();
        bt_scpanel.setBounds(20, 20, 470, 187);
        getContentPane().add(bt_scpanel);

		/* 2003.10.20 tuika y.k */
        mabiki_button = new JButton("�Ԉ���");
        mabiki_button.setBounds(20, 237, 100, 24);
        mabiki_button.setLocale(new Locale("ja","JP"));
        mabiki_button.setFont(new java.awt.Font("dialog", 0, 18));
        mabiki_button.setBorder(new Flush3DBorder());
        mabiki_button.setForeground(java.awt.Color.black);
        mabiki_button.addActionListener(new MabikiButton());
        mabiki_button.setEnabled(true);
        getContentPane().add(mabiki_button);

        JLabel  lab2 = new JLabel("�ΐF�F���Ԉ����@���F�F�Ԉ����ς݁@���F�F�đ����@�s���N�F�đ����@�F�F�Ԉ����w���ς�");
        lab2.setBounds(20, 207, 600, 24);
        lab2.setLocale(new Locale("ja","JP"));
        lab2.setFont(new java.awt.Font("dialog", 0, 12));
//        lab2.setBorder(new Flush3DBorder());
        lab2.setForeground(java.awt.Color.black);
//        lab2.setBackground(java.awt.Color.lightGray);
        getContentPane().add(lab2);

//        lab2 = new JLabel("�s���N�F�đ����@�F�F�Ԉ����w���ς�",JLabel.CENTER);
//        lab2.setBounds(140, 251, 300, 24);
//        lab2.setLocale(new Locale("ja","JP"));
//        lab2.setFont(new java.awt.Font("dialog", 0, 12));
//        lab2.setBorder(new Flush3DBorder());
//        lab2.setForeground(java.awt.Color.black);
//        getContentPane().add(lab2);
		/* 2003.10.20 tuika end y.k */

		/* �E��i�̃p�l�� */
        bt_start_scpanel = new JScrollPane();
        bt_start_scpanel.setBounds(510, 20, 410, 187);
        getContentPane().add(bt_start_scpanel);

		/* �����i�̃p�l�� */
        bt_condition_scpanel = new JScrollPane();
        bt_condition_scpanel.setBounds(20, 227+iIdo, 860, 187);
        getContentPane().add(bt_condition_scpanel);

        pv_panel = new JPanel();
        pv_panel.setLayout(null);
        pv_panel.setBounds(20, 434+iIdo, 860, 230);
        pv_panel.setBorder(new Flush3DBorder());
        pv_panel.setBackground(java.awt.Color.lightGray);
        getContentPane().add(pv_panel);

		/* ���i�̃p�l�� */
        pv_table = new PVTable();
        JScrollPane panel = new JScrollPane(pv_table);
        panel.setBounds(20, 20, 488 ,187);
        pv_panel.add(panel);

        unit_cbox = new JComboBox();
        unit_cbox.setBounds(520, 20, 80, 24);
        unit_cbox.setLocale(new Locale("ja","JP"));
        unit_cbox.setFont(new java.awt.Font("dialog", 0, 12));
        unit_cbox.setForeground(java.awt.Color.black);
        unit_cbox.addItem("����");
        unit_cbox.addItem("����");
		unit_cbox.setFocusable(false);	/* 2007.08.22 */
        unit_cbox.addActionListener(new ChgUnit());
        pv_panel.add(unit_cbox);

        JLabel  lab = new JLabel("�J�n",JLabel.CENTER);
        lab.setBounds(600, 20, 50, 24);
        lab.setLocale(new Locale("ja","JP"));
        lab.setFont(new java.awt.Font("dialog", 0, 12));
        lab.setBorder(new Flush3DBorder());
        lab.setForeground(java.awt.Color.black);
        pv_panel.add(lab);

        lab = new JLabel("�I��",JLabel.CENTER);
        lab.setBounds(600, 50, 50, 24);
        lab.setLocale(new Locale("ja","JP"));
        lab.setFont(new java.awt.Font("dialog", 0, 12));
        lab.setBorder(new Flush3DBorder());
        lab.setForeground(java.awt.Color.black);
        pv_panel.add(lab);

        lab = new JLabel("�Ԋu",JLabel.CENTER);
        lab.setBounds(600, 80, 50, 24);
        lab.setLocale(new Locale("ja","JP"));
        lab.setFont(new java.awt.Font("dialog", 0, 12));
        lab.setBorder(new Flush3DBorder());
        lab.setForeground(java.awt.Color.black);
        pv_panel.add(lab);

        start_text = new NumeralText(6,false);
        start_text.setBounds(650, 20, 80, 24);
        start_text.setText("0");
        pv_panel.add(start_text);

        end_text = new NumeralText(6,false);
        end_text.setBounds(650, 50, 80, 24);
        end_text.setText("0");
        pv_panel.add(end_text);

        interval_text = new NumeralText(6,false);
        interval_text.setBounds(650, 80, 80, 24);
        interval_text.setText("0");
        pv_panel.add(interval_text);

        start_lab = new JLabel("X 10s",JLabel.CENTER);
        start_lab.setBounds(730, 20, 60, 24);
        start_lab.setLocale(new Locale("ja","JP"));
        start_lab.setFont(new java.awt.Font("dialog", 0, 12));
        start_lab.setBorder(new Flush3DBorder());
        start_lab.setForeground(java.awt.Color.black);
        pv_panel.add(start_lab);

        end_lab = new JLabel("X 10s",JLabel.CENTER);
        end_lab.setBounds(730, 50, 60, 24);
        end_lab.setLocale(new Locale("ja","JP"));
        end_lab.setFont(new java.awt.Font("dialog", 0, 12));
        end_lab.setBorder(new Flush3DBorder());
        end_lab.setForeground(java.awt.Color.black);
        pv_panel.add(end_lab);

        interval_lab = new JLabel("X 10s",JLabel.CENTER);
        interval_lab.setBounds(730, 80, 60, 24);
        interval_lab.setLocale(new Locale("ja","JP"));
        interval_lab.setFont(new java.awt.Font("dialog", 0, 12));
        interval_lab.setBorder(new Flush3DBorder());
        interval_lab.setForeground(java.awt.Color.black);
        pv_panel.add(interval_lab);

        JTextField hed = new JTextField(save_dir);
        hed.setBounds(520, 125, 270, 24);
        hed.setEnabled(false);
        pv_panel.add(hed);

        file_text = new FileText();
        file_text.setBounds(520, 150, 270, 24);
        pv_panel.add(file_text);

        save_button = new JButton("���ѕۑ�");
        save_button.setBounds(520, 183, 100, 24);
        save_button.setLocale(new Locale("ja","JP"));
        save_button.setFont(new java.awt.Font("dialog", 0, 18));
        save_button.setBorder(new Flush3DBorder());
        save_button.setForeground(java.awt.Color.black);
        save_button.addActionListener(new SaveButton());
        save_button.setEnabled(false);
        pv_panel.add(save_button);

        // ======================================== [�ݒ�Ǎ�]�{�^�� ==================================
        set_read_button = new JButton("�ݒ�Ǎ�");
        set_read_button.setBounds(630, 183, 100, 24);
        set_read_button.setLocale(new Locale("ja","JP"));
        set_read_button.setFont(new java.awt.Font("dialog", 0, 18));
        set_read_button.setBorder(new Flush3DBorder());
        set_read_button.setForeground(java.awt.Color.black);
        set_read_button.addActionListener(
          new ActionListener() {
              public void actionPerformed(ActionEvent evt) {
                  JFileChooser chooser = new JFileChooser(file_);
                  int ret = chooser.showOpenDialog(pvdatasave);
                  if ( ret == JFileChooser.APPROVE_OPTION ) {
                      file_ = chooser.getSelectedFile();        // �t�@�C�������擾����
                      Properties prop = new Properties();       // �v���p�e�B�𐶐�����
                      try {
                          FileInputStream in = new FileInputStream(file_);
                          prop.load( in );                      //�v���p�e�B���擾����B
                          in.close();
                          prop.list(System.out);

                          // �\���ݒ�̓Ǎ�
                          prop_Unit     = prop.getProperty("SELUNIT");
                          prop_Start    = prop.getProperty("START");
                          prop_End      = prop.getProperty("END");
                          prop_Interval = prop.getProperty("INTERVAL");

                          // PV���ڂ̓Ǎ�
                          for(int i=0; i < 128 ; i++){
                              try {
                                  prop_PVItemNo[i]   = prop.getProperty("PVDATA" + (i+1));
                              } catch (Exception e) {
                                  prop_PVItemNo[i]   = new String("");
                              }
                          }

                          pv_table.clearSelection(); // �SPV���ڑI���N���A
                          // �t�@�C������ǂݍ���PV���ڂ�I��
                          for(int i = 0; i < 128 ; i++){
                              if(prop_PVItemNo[i].equals("1")){
                                  CZSystem.log("CZPVDataSave","PVDATA" + i + " : " + prop_PVItemNo[i]);
                                  pv_table.addRowSelectionInterval(i,i);
                              }
                          }

                          // ���� or ���� �ݒ�
                          if(prop_Unit.equals("1")){
                              unit_cbox.setSelectedIndex(1);
                          }else{
                              unit_cbox.setSelectedIndex(0);
                          }

                          // �J�n�E�I���E�Ԋu �ݒ�
                          start_text.setText(prop_Start);
                          end_text.setText(prop_End);
                          interval_text.setText(prop_Interval);

                      } catch ( IOException ex ) {
                          CZSystem.log("CZPVDataSave","PVITEM Property File�����[�h�ł��܂���ł����B");
                          return;
                      }
                  }
              }
          }
        );
        pv_panel.add(set_read_button);

        // ======================================== [�ݒ�ۑ�]�{�^�� ==================================
        set_save_button = new JButton("�ݒ�ۑ�");
        set_save_button.setBounds(740, 183, 100, 24);
        set_save_button.setLocale(new Locale("ja","JP"));
        set_save_button.setFont(new java.awt.Font("dialog", 0, 18));
        set_save_button.setBorder(new Flush3DBorder());
        set_save_button.setForeground(java.awt.Color.black);
        set_save_button.addActionListener(
          new ActionListener() {
              public void actionPerformed(ActionEvent evt)
              {
                  JFileChooser chooser = new JFileChooser(file_);
                  int ret = chooser.showSaveDialog(pvdatasave);
                  if (ret == JFileChooser.APPROVE_OPTION) {
                      file_ = chooser.getSelectedFile();            // �t�@�C�������擾����
                      Properties prop = new Properties();           // �v���p�e�B�𐶐�����
                      // �\���̐ݒ�
                      prop.setProperty(new String("SELUNIT"),  new String("" + unit_cbox.getSelectedIndex()) );
                      prop.setProperty(new String("START"),    new String("" + start_text.getText()));
                      prop.setProperty(new String("END"),      new String("" + end_text.getText())  );
                      prop.setProperty(new String("INTERVAL"), new String("" + interval_text.getText()) );

                      // PV���ڂ̐ݒ�
                      for (int i = 0; i < 128; i++) {
                        if(pv_table.isRowSelected(i) == true){
                            prop.setProperty(new String("PVDATA" + (i+1)),  new String("1"));
                        }else{
                            prop.setProperty(new String("PVDATA" + (i+1)),  new String("0"));
                        }
                      }
                      //---------- �t�@�C���ɕۑ�����  ----------
                      try {
//						CZSystem.log("CZTPGFrame ","�t�@�C���ɕۑ������B");
                          FileOutputStream out = new FileOutputStream(file_);
                          prop.store(out, "");
                          out.flush();
                          out.close();
                      } catch (IOException ex) {
                          JOptionPane.showMessageDialog(
                            pvdatasave,
                            new String("�ۑ��ł��܂���ł����B"),
                            new String("�ۑ�"),
                            JOptionPane.WARNING_MESSAGE);
                          return;
                      }
                      JOptionPane.showMessageDialog(
                        pvdatasave,
                        new String("�ۑ����܂����B"),
                        new String("�ۑ�"),
                        JOptionPane.INFORMATION_MESSAGE);
                      return;
                  }
              }
          }
        );
        pv_panel.add(set_save_button);


    }

    //
    // ��ʏ����\��
    public boolean setDefault(){

        clearFileName();            // �ۑ�����File�����N���A����B
        removeBtCondition();        // �o�b�`�����폜����B
/* 2003.10.20 y.k */
        //�o�b�`�����쐬����B
//        BtTable t = new BtTable();
       	 btTable = new BtTable();

        JTableHeader tabHead = btTable.getTableHeader();
        tabHead.setReorderingAllowed(false);
        bt_scpanel.setViewportView(btTable);
        // �ۑ��{�^����L���ɂ���B
        save_button.setEnabled(false);
        // ����n�Q�Ƌ@�\    @20131021
        if(CZSystemDefine.REFERENCE_RUN == CZSystem.getRunLevel()){
            mabiki_button.setEnabled(false);
        }
        return true;
    }

    //
    // �J�[�\����ݒ肷��B
    private void setCur(Cursor cu){
        setCursor(cu);
    }

    //
    // �J�[�\�����擾����B
    private Cursor getCur(){
        return getCursor();
    }   

    //
    // File�����N���A����B
    public void clearFileName(){
//@@        CZSystem.log("CZPVDataSave","clearFileName");
        String text = "";
        file_text.setText(text);
        // �ۑ��{�^���𖳌��ɂ���B
        save_button.setEnabled(false);
        return;
    }

    //
    // File����ݒ肷��B
    public void setFileName(int row){
        // �s��I�����Ă��Ȃ����́AFile�����N���A����B
        if(0 > row){
            clearFileName();
            return;
        }
        // �J�n��񂪖������́AFile�����N���A����B
        CZSystemStart st = (CZSystemStart)start_list.elementAt(row);
        if(null == st){
            clearFileName();
            return;
        }
        // �J�n�����File�����쐬����B
//@@        CZSystem.log("CZPVDataSave setFileName()",
//@@            "savePVData [" + st.batch + "][" + st.p_no + "][" + st.p_renban + "][" + st.p_start + "]");
        String file_name = new String( st.batch.trim() + "-" +
                       st.p_no         + "-" +
                       st.p_renban     + ".csv");
//@@        CZSystem.log("CZPVDataSave setFileName()","setFileName [" + row + "][" + file_name + "]");
        file_text.setText(file_name);
        // �ۑ��{�^����L���ɂ���B
        save_button.setEnabled(true);
        return;
    }


    //
    //
    // �o�b�`�����쐬����B
    public boolean setBtCondition(Vector v){

//@@        CZSystem.log("CZPVDataSave","setBtCondition() Start");
        removeBtCondition();

        BtStartTable t = new BtStartTable(v);
        JTableHeader tabHead = t.getTableHeader();
        tabHead.setReorderingAllowed(false);
        bt_start_scpanel.setViewportView(t);

        BtConditionTable tt = new BtConditionTable(v);
        tabHead = tt.getTableHeader();
        tabHead.setReorderingAllowed(false);
        bt_condition_scpanel.setViewportView(tt);

        return true;
    }

    //
    // �o�b�`�����폜����B
    public boolean removeBtCondition(){

        JViewport v;
        v =  bt_start_scpanel.getViewport();
        if(null != v.getView()) v.remove(v.getView());

        v =  bt_condition_scpanel.getViewport();
        if(null != v.getView()) v.remove(v.getView());
        return true;
    }

    //
    // �o�b�`�J�n����ێ�����B
    public void setBtStartList(Vector v){
        start_list = v;
    }

    //
    // PV�f�[�^��ۑ�����B
    public boolean savePVData(int row){
        // File�����ݒ肳��Ă��Ȃ����̓G���[�ɂ���B
        String file_name = file_text.getText();
        if(1 > file_name.length()) return false;
        // �o�b�`�J�n��񂪖������̓G���[�ɂ���B
        if(null == start_list) return false;
        if(0 > row) return false;

//@@        CZSystem.log("CZPVDataSave savePVData()"," [" + row + "]");
        // �L���ȊJ�n����I�����Ă��Ȃ����̓G���[�ɂ���B
        CZSystemStart st = (CZSystemStart)start_list.elementAt(row);
        if(null == st) return false;

//@@        CZSystem.log("CZPVDataSave savePVData()",
//@@            " [" + st.batch + "][" + st.p_no + "][" + st.p_renban + "][" + st.p_start + "]");
        // PV�f�[�^��ێ�����̈���m�ۂ��A�N���A����B
        boolean data_no[] = new boolean[CZSystemDefine.PV_MAX_LENGTH];
        for(int i = 0 ; i < CZSystemDefine.PV_MAX_LENGTH ; i++) data_no[i] = false;

        int list[];
        int no;
        list = pv_table.getSelectedRows();
        if(1 > list.length) return false;

        for(int i = 0 ; i < list.length ; i++){
            CZSystem.log("CZPVDataSave savePVData()","savePVData List[" + i + "][" + list[i] + "]");
    
            no = list[i];
            if(no < CZSystemDefine.PV_MAX_LENGTH) data_no[no] = true;
        }

        String db = CZSystem.getDBName();
        String view = CZSystem.getViewName(db,st.batch);
    
        Vector pv_data = CZSystem.getPVData(db,view,st.p_renban,data_no);
        if(null == pv_data) return false;
        if(1 > pv_data.size()) return false;
        String jhed = createJHed(data_no);
        String hed = createHed(data_no);
//@@        CZSystem.log("CZPVDataSave savePVData()","Head [" + hed + "]" );

        writeFile(jhed,hed,pv_data,data_no);

        return true;
    }


    //
    //
    //
    private int writeFile(String jhed,String hed,Vector pv_data,boolean data_no[]){
        int write_count = 1;

        File file = new File(save_dir,file_text.getText());
        FileOutputStream s = null;
        PrintWriter pr     = null;

        float pos = 0;
        float next_pos = 0;

        int start = 0;
        int end   = 0;
        int inc   = 0;

        start = start_text.getValue();  
        end   = end_text.getValue();    
        inc   = interval_text.getValue();   

        int index = unit_cbox.getSelectedIndex();
        switch(index){
            case 0:
            default:
                start = start * 10; 
                end   = end * 10;   
                inc   = inc * 10;   
                break;
            case 1:
                break;
        }

        try{

            s = new FileOutputStream(file);
            pr = new PrintWriter(s);
            // Header���������B
            pr.println(jhed);
            // Header���������B
            pr.println(hed);
            // PV�f�[�^�������B
            for(int i = 0 ; i < pv_data.size() ; i++){
                CZSystemPVData pv = (CZSystemPVData)pv_data.elementAt(i);

                switch(index){
                case 0:
                default:
                    pos = pv.p_time;
                    break;
                case 1:
                    pos = pv.p_length;  
                    break;
                }

                if(start  > pos) continue;
                if((end != 0) && (end < pos)) break;
                if(1 == write_count) next_pos = start;

                if(next_pos <= pos){
                    writeRec(pr,pv,data_no);
                    write_count++;
                    next_pos += inc;
                }
            }
        }
        catch(IOException e){
            CZSystem.log("CZPVDataSave writeFile()","[" + write_count + "][" + e + "]");
            if(null != pr) pr.close();
            return -1;
        }

        if(null != pr) pr.close();

        write_count--;
        return write_count;
    }

    //
    // PV�f�[�^��CSV�`���ŏ����B
    private boolean writeRec(PrintWriter pr,CZSystemPVData pv,boolean data_no[]){

        StringBuffer rec = new StringBuffer(3000);  

        try{
            rec.append(pv.p_no);
            rec.append("," + pv.sp_no);
            rec.append("," + pv.p_renban);
            rec.append("," + pv.p_time);
            rec.append("," + pv.sp_time);
            rec.append("," + pv.p_date);
            rec.append("," + pv.h_ontime);
            rec.append("," + pv.hk_renban);

            for(int j = 0 ; j < data_no.length ; j++){
                if(!data_no[j]) continue;
                rec.append("," + pv.data[j]);
            }

            pr.println(rec);
        }
        catch(Exception e){
            CZSystem.log("CZPVDataSave writeRec()","[" + e + "]");
            return false;
        }
        return true;
    }

    //
    // Header��1���쐬����B���{�ꖼ�̕�
    private String createJHed(boolean v[]){

        StringBuffer buf = new StringBuffer(2000);

        buf.append("�v���Z�XNo,�T�u�v���Z�XNo,�v���Z�X�A��,�v���Z�X����,�T�u�v���Z�X����,�̎����,���C���q�[�^�d���I������,���グ�������A��");
        for(int i = 0 ; i < v.length ; i++){
            if(!v[i]) continue;
            CZSystemPVName name = CZSystem.getPVName(i);
            if(null == name) break;
            buf.append("," + name.j_name.trim());
        }
        String ret = buf.toString();
        return ret;
    }

    //
    // Header��2���쐬����B
    private String createHed(boolean v[]){

        StringBuffer buf = new StringBuffer(2000);

        buf.append("p_no,sp_no,p_renban,p_time,sp_time,p_date,h_ontime,hk_renban");
        for(int i = 0 ; i < v.length ; i++){
            if(!v[i]) continue;
            CZSystemPVName name = CZSystem.getPVName(i);
            if(null == name) break;
            buf.append("," + name.k_name.trim());
        }
        String ret = buf.toString();
        return ret;
    }

    /***************************************************************************
     *  ����or�����I��ComboBox�̃��X�i�[
     *  (�I���ɍ��킹�ĒP�ʕ\����ς���) X 10s <--> mm  
     ***************************************************************************/
    class ChgUnit implements ActionListener {
        public void actionPerformed(ActionEvent ev){

            int index = unit_cbox.getSelectedIndex();

//@@            CZSystem.log("CZPVDataSave ChgUnit()","[" + index + "]" );

            switch(index){
            case 0:
            default:
                start_lab.setText("X 10s");
                end_lab.setText("X 10s");
                interval_lab.setText("X 10s");
                break;
            case 1:
                start_lab.setText("mm");
                end_lab.setText("mm");
                interval_lab.setText("mm");
                break;
            }
        }
    }

    /***************************************************************************
     *
     *  �ۑ��{�^���̃��X�i�[
     *
     ***************************************************************************/
    class SaveButton implements ActionListener {
        public void actionPerformed(ActionEvent ev){

            JViewport v;
            JTable t;

            v = bt_scpanel.getViewport();
            t = (JTable)v.getView();
            if(null == t) return;
            int bt_row = t.getSelectedRow();

            v = bt_start_scpanel.getViewport();
            t = (JTable)v.getView();
            if(null == t) return;
            int bt_start_row = t.getSelectedRow();
//@@            CZSystem.log("CZPVDataSave SaveButton","valueChanged [" + bt_row + "][" + bt_start_row + "]");

            // �J�[�\����Wait�ɐݒ肷��B
            Cursor cu_tmp = getCur();
            Cursor cu = new Cursor(Cursor.WAIT_CURSOR);
            setCur(cu);
            // PV�f�[�^��ۑ�����B
            savePVData(bt_start_row);
            // �J�[�\�������ɖ߂��B
            setCur(cu_tmp);

        }
    }

    /***************************************************************************
     *
     *  �Ԉ����{�^���̃��X�i�[�@2003.10.20 y.k tuika 
     *
     ***************************************************************************/
	@SuppressWarnings("unchecked")
    class MabikiButton implements ActionListener {
        public void actionPerformed(ActionEvent ev){
	        Vector  cngBtColor_list   = null;
			int	iCnt = 0;

			cngBtColor_list = new Vector(100);

			/* �{�^�����������Ƃ��̏��������ɒǉ� */
            Object msg2[] = {"�đ��o�u���т̊Ԉ��������̎w�����s���܂��B��낵���ł����H"};

            if(confirmDia(msg2))
			{
				System.out.println ("MabikiButton�F���s");
				for (int iLp=0; iLp < btTable.dispBtColor_list.size(); iLp++)
				{
					DispBtColorTbl bt = (DispBtColorTbl)btTable.dispBtColor_list.elementAt(iLp);
//					if ((bt.m_sumi == 3) && (bt.m_sumi_chg == -1))
					if ((bt.m_sumi == 3) && (bt.m_sumi_chg == bt.m_sumi))
//					if (bt.m_sumi == 3)
					{	/* �đ������f�[�^���Ԉ����w���ς� */
						bt.m_sumi_chg = 4;
						cngBtColor_list.addElement(bt);
						iCnt++;
					}
				}
//System.out.println ("iCnt:" + iCnt + "count:" + cngBtColor_list.size());
				if (iCnt > 0)
					CZSystem.CZPvControlChgSend("MABIKI",CZSystem.getRoName(), cngBtColor_list);
				else
				{
			        JOptionPane.showMessageDialog(null,"�Ԉ����ł���o�b�`�͂���܂���",
                                    "�đ��Ԉ�������",JOptionPane.ERROR_MESSAGE);
				}
//		        mabiki_button.setEnabled(false);
			}
			else
			{
				System.out.println ("MabikiButton�F�L�����Z��");
			}
        }
    }

        //
        // �m�F���b�Z�[�W�̕\��
        //
        private boolean confirmDia(Object msg[]){
	            int ans = JOptionPane.showConfirmDialog(null,msg,
	                    "�đ��o�u���ъԈ������s�m�F�_�C�A���O",
	                    JOptionPane.OK_CANCEL_OPTION,
	                    JOptionPane.WARNING_MESSAGE);
	            if(0 == ans) return true;
	            return false;
        }

    /****************************************************************************
     *
     *       �a���m���ꗗ
     *
     ****************************************************************************/
    class BtTable extends JTable {

        private Vector  bt_all_list = null;
        private Vector  bt_list     = null;
        private Vector  pvControl_all_list = null;	/* 2003.10.21 y.k */
        private Vector  dispBtColor_list   = null;
		private	CZSystemPvControl pvcDt = null;
        private BtTblMdl model = null;
        
        private boolean life = false;
		private int iLp;

		@SuppressWarnings("unchecked")
        BtTable(){
            super();

//System.out.println ("BtTable : start");
            try{
                setName("BtTable");
                setAutoCreateColumnsFromModel(true);
                setSelectionMode(ListSelectionModel.SINGLE_SELECTION);

                setLocale(new Locale("ja","JP"));
                setFont(new java.awt.Font("dialog", 0, 12));
                setRowHeight(17);

//@@                CZSystem.log("CZPVDataSave BtTable ","new");

                bt_all_list = CZSystem.getBtCondition(CZSystem.getDBName());

                if(null == bt_all_list) return;

			    //  ���Ƃo�u���ъǗ����擾
				pvControl_all_list = CZSystem.getPvControl(CZSystem.getDBName());

                bt_list = new Vector(100);
				dispBtColor_list = new Vector(100);
                for(int i = 0 ; i < bt_all_list.size() ; i++){
                    CZSystemBt bt = (CZSystemBt)bt_all_list.elementAt(i);

                    if(0 == bt.renban)
					{
						bt_list.addElement(bt);
						/* 2003.10.21 tuika y.k */
						DispBtColorTbl dispBtColor = new DispBtColorTbl();
						dispBtColor.batch = bt.batch;
						dispBtColor.t_name = null;
						dispBtColor.m_flg = -1;
						dispBtColor.m_sumi = -1;
						dispBtColor.m_sumi_chg = dispBtColor.m_sumi;
						if ((pvControl_all_list != null) && (pvControl_all_list.size() > 0))
						{
							for (iLp = 0 ; iLp < pvControl_all_list.size() ; iLp++)
							{
			                    pvcDt = (CZSystemPvControl)pvControl_all_list.elementAt(iLp);
//	System.out.println ("batch [" + dispBtColor.batch + ":" + dispBtColor.batch.length() + "][" + pvcDt.batch + ":" + pvcDt.batch.length() + "]");

								if (dispBtColor.batch.equals(pvcDt.batch))
								{
//	System.out.println ("batch Data set");
//									dispBtColor.t_name = pvcDt.t_name;
//									dispBtColor.m_flg = pvcDt.m_flg;
									dispBtColor.m_sumi = pvcDt.m_sumi;
									dispBtColor.m_sumi_chg = dispBtColor.m_sumi;
									break;
								}
							}
						}
						dispBtColor_list.addElement(dispBtColor);
					}
//@@@@@@@                    if(-1 == bt.renban) bt_list.addElement(bt);
                }

//@@                CZSystem.log("CZPVDataSave BtTable ","bt_list OK");

                model = new BtTblMdl(bt_list);
                setModel(model);

//@@                CZSystem.log("CZPVDataSave BtTable ","model OK");

                DefaultTableColumnModel cmdl = (DefaultTableColumnModel)getColumnModel();
                TableColumn colum = null;
/* 2003.10.20 y.k */
	            ColorRender ren   = null;


                // No
	            ren = new ColorRender();
//            ren.setHorizontalAlignment(ren.CENTER);

                colum = cmdl.getColumn(0);
                colum.setMaxWidth(40);
                colum.setMinWidth(40);
                colum.setWidth(40);
	            colum.setCellRenderer(ren);

                // BtNo
	            ren = new ColorRender();
                colum = cmdl.getColumn(1);
                colum.setMaxWidth(130);
                colum.setMinWidth(130);
                colum.setWidth(130);
	            colum.setCellRenderer(ren);

                // �i��
                ren = new ColorRender();
                colum = cmdl.getColumn(2);
                colum.setMaxWidth(80);
                colum.setMinWidth(80);
                colum.setWidth(80);
                colum.setCellRenderer(ren);

                // T2
                ren = new ColorRender();
                colum = cmdl.getColumn(3);
                colum.setMaxWidth(40);
                colum.setMinWidth(40);
                colum.setWidth(40);
                colum.setCellRenderer(ren);

                // �o�^����
                ren = new ColorRender();
                colum = cmdl.getColumn(4);
                colum.setMaxWidth(162);
                colum.setMinWidth(162);
                colum.setWidth(162);
                colum.setCellRenderer(ren);
            }
            catch (Throwable e) {
                CZSystem.handleException(e);
            }

//System.out.println ("BtTable : end");

        }

        //
        //
        //
		@SuppressWarnings("unchecked")
        public void valueChanged(ListSelectionEvent e){
//System.out.println ("valueChanged : start");

            super.valueChanged(e);

            if(e.getValueIsAdjusting()) return;

            int row = getSelectedRow();

//@@            CZSystem.log("CZPVDataSave SelectionEvent",
//@@                "valueChanged [" + row + "][" + getSelectedColumn() + "]");

            if(0 > row){
                if(!life){
                    life = true;
                    return;
                }
                clearFileName();
                removeBtCondition();
                return;
            }

            clearFileName();

            Vector v = new Vector(50);

            CZSystemBt bt = (CZSystemBt)bt_list.elementAt(row);

            for(int i = 0 ; i < bt_all_list.size() ; i++){
                CZSystemBt bt_tmp = (CZSystemBt)bt_all_list.elementAt(i);

                if(bt.batch.equals(bt_tmp.batch)) v.addElement(bt_tmp);
            }

            setBtCondition(v);
        }

        //
        //
        //
        public void setData(int gr,int tbl){
//@@            CZSystem.log("CZPVDataSave setData()","[" + gr + "][" + tbl + "]");
        }

    }

        /**
        *
        */
        class ColorRender extends DefaultTableCellRenderer {

            ColorRender(){
                super();
            }

            public Component getTableCellRendererComponent( JTable table,
                                                        Object value,
                                                        boolean isSelected,
                                                        boolean hasFocus,
                                                        int row,int column){

				if (btTable != null)
				{
					if (btTable.dispBtColor_list != null)
					{
						DispBtColorTbl bt = (DispBtColorTbl)btTable.dispBtColor_list.elementAt(row);
						if (bt.m_sumi == 0)			/* ���Ԉ��� */
			                setBackground(java.awt.Color.green);
						else if (bt.m_sumi == 1)		/* �Ԉ����ς� */
			                setBackground(java.awt.Color.white);
						else if (bt.m_sumi == 2)		/* �đ��� */
			                setBackground(java.awt.Color.yellow);
						else if (bt.m_sumi == 3)		/* �đ��� */
			                setBackground(java.awt.Color.pink);
						else if (bt.m_sumi == 4)		/* �Ԉ����w���ς� */
//			                setBackground(java.awt.Color.blue);
			                setBackground(java.awt.Color.cyan);
						else if (bt.m_sumi == -1)	/* �Y���a�������� */
			                setBackground(java.awt.Color.green);
						else
			                setBackground(java.awt.Color.lightGray);
					}
				}

                super.getTableCellRendererComponent(table,
                                                    value,
                                                    isSelected,
                                                    hasFocus,
                                                    row,column);
                return(this);
            }
        } //class ColorRender extends DefaultTableCellRenderer

    /****************************************************************************
     *
     *       �a���m�����шꗗ�F���f��
     *
     ****************************************************************************/
    public class BtTblMdl extends AbstractTableModel {

        private int TBL_ROW     = 0;
        final   int TBL_COL     = 5;
        private Vector  bt_list     = null;

        final String[] names = {" # "  , "Bt" , "�i��" , "T2" , "�o�^����" };

        private Object  data[][];

        BtTblMdl(Vector v){
            super();

            bt_list = v;
            TBL_ROW = bt_list.size();

            data = new Object[TBL_ROW][TBL_COL];

//@@            CZSystem.log("CZPVDataSave BtTblMdl ","new size[" + TBL_ROW + "]");

            for(int i = 0 ; i < TBL_ROW ; i++){
                CZSystemBt bt = (CZSystemBt)bt_list.elementAt(i);

                if(null == bt){
                    CZSystem.log("CZPVDataSave BtTblMdl ","bt[null]");
                    break;
                }

                data[i][0] = new Integer(i+1);
                data[i][1] = bt.batch;
                data[i][2] = bt.hinshu;
                data[i][3] = bt.no_hikiage;
                data[i][4] = bt.t_time;
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


    /****************************************************************************
     *
     *       �a���X�^�[�g���Ԉꗗ
     *
     ****************************************************************************/
    class BtStartTable extends JTable {

        private Vector  bt_list         = null;
        private Vector  bt_start_list   = null;

        private BtStartTblMdl model     = null;

        private boolean life            = false;

        BtStartTable(Vector v){
            super();
            bt_list = v;

            try{
                setName("BtStartTable");
                setAutoCreateColumnsFromModel(true);
                setSelectionMode(ListSelectionModel.SINGLE_SELECTION);

                setLocale(new Locale("ja","JP"));
                setFont(new java.awt.Font("dialog", 0, 12));
                setRowHeight(17);

                CZSystemBt bt = (CZSystemBt)bt_list.elementAt(0);
                bt_start_list = new Vector(100);
                bt_start_list = CZSystem.getBtStart(CZSystem.getDBName(),bt.batch);
                setBtStartList(bt_start_list);

                //NULL���K�v
                if(null == bt_start_list) return;

                model = new BtStartTblMdl(bt_start_list);
                setModel(model);

                DefaultTableColumnModel cmdl = (DefaultTableColumnModel)getColumnModel();
                TableColumn colum = null;

                // No
                colum = cmdl.getColumn(0);
                colum.setMaxWidth(40);
                colum.setMinWidth(40);
                colum.setWidth(40);

                // PNo
                colum = cmdl.getColumn(1);
                colum.setMaxWidth(40);
                colum.setMinWidth(40);
                colum.setWidth(40);

                // SPNo
                colum = cmdl.getColumn(2);
                colum.setMaxWidth(40);
                colum.setMinWidth(40);
                colum.setWidth(40);

                // PSeq
                colum = cmdl.getColumn(3);
                colum.setMaxWidth(40);
                colum.setMinWidth(40);
                colum.setWidth(40);

                // �v���Z�X
                colum = cmdl.getColumn(4);
                colum.setMaxWidth(70);
                colum.setMinWidth(70);
                colum.setWidth(70);

                // �o�^����
                colum = cmdl.getColumn(5);
                colum.setMaxWidth(162);
                colum.setMinWidth(162);
                colum.setWidth(162);
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

            if(e.getValueIsAdjusting()) return;
        
            int row = getSelectedRow();

//@@            CZSystem.log("CZPVDataSave ",
//@@                "valueChanged [" + row + "][" + getSelectedColumn() + "]");

            if(0 > row){
                if(!life){
                    life = true;
                    return;
                }
                clearFileName();
                return;
            }

            setFileName(row);
        }

        //
        //
        //
        public void setData(int gr,int tbl){
//@@            CZSystem.log("CZPVDataSave setData()","[" + gr + "][" + tbl + "]");
        }
    }

    /****************************************************************************
     *
     *       �a���X�^�[�g���Ԉꗗ�F���f��
     *
     ****************************************************************************/
    public class BtStartTblMdl extends AbstractTableModel {

        private int TBL_ROW     = 0;
        final   int TBL_COL     = 6;
        private Vector  bt_start_list   = null;

        final String[] names = {" # "  , "PNo" ,    
                    "SPNo","PSeq"  ,
                    "�v���Z�X",
//                    "�o�^����" };
                    "�J�n����" };

        private Object  data[][];

        BtStartTblMdl(Vector v){
            super();
            bt_start_list = v;
            TBL_ROW = bt_start_list.size();

            data = new Object[TBL_ROW][TBL_COL];

            for(int i = 0 ; i < TBL_ROW ; i++){
                CZSystemStart st = (CZSystemStart)bt_start_list.elementAt(i);

                if(null == st) break;

                data[i][0] = new Integer(i+1);
                data[i][1] = new Integer(st.p_no);
                data[i][2] = new Integer(st.sp_no);
                data[i][3] = new Integer(st.p_renban);
                data[i][4] = CZSystem.getProcName(st.p_no);
                data[i][5] = st.p_start;
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


    /***************************************************************************
     *
     *       �a���o�^���ꗗ
     *
     ****************************************************************************/
    class BtConditionTable extends JTable {

        private Vector  bt_list     = null;

        private BtConditionTblMdl model = null;

        BtConditionTable(Vector v){
            super();
            bt_list = v;

            try{
                setName("BtConditionTable");
                setAutoCreateColumnsFromModel(true);
                setSelectionMode(ListSelectionModel.SINGLE_SELECTION);

                setLocale(new Locale("ja","JP"));
                setFont(new java.awt.Font("dialog", 0, 12));
                setRowHeight(17);

                model = new BtConditionTblMdl(bt_list);
                setModel(model);

                DefaultTableColumnModel cmdl = (DefaultTableColumnModel)getColumnModel();
                TableColumn colum = null;

                // No
                colum = cmdl.getColumn(0);
                colum.setMaxWidth(40);
                colum.setMinWidth(40);
                colum.setWidth(40);

                // �o�^����
                colum = cmdl.getColumn(1);
                colum.setMaxWidth(140);
                colum.setMinWidth(140);
                colum.setWidth(140);

                // �A��
                colum = cmdl.getColumn(2);
                colum.setMaxWidth(40);
                colum.setMinWidth(40);
                colum.setWidth(40);

                // �i��
                colum = cmdl.getColumn(3);
                colum.setMaxWidth(80);
                colum.setMinWidth(80);
                colum.setWidth(80);

                // ���c�{
                colum = cmdl.getColumn(4);
                colum.setMaxWidth(40);
                colum.setMinWidth(40);
                colum.setWidth(40);

                // ���a
                colum = cmdl.getColumn(5);
                colum.setMaxWidth(40);
                colum.setMinWidth(40);
                colum.setWidth(40);

                // ���㒷
                colum = cmdl.getColumn(6);
                colum.setMaxWidth(40);
                colum.setMinWidth(40);
                colum.setWidth(40);

                // ���d��
                colum = cmdl.getColumn(7);
                colum.setMaxWidth(60);
                colum.setMinWidth(60);
                colum.setWidth(60);

                // �ǎd��
                colum = cmdl.getColumn(8);
                colum.setMaxWidth(60);
                colum.setMinWidth(60);
                colum.setWidth(60);

                // T1
                colum = cmdl.getColumn(9);
                colum.setMaxWidth(40);
                colum.setMinWidth(40);
                colum.setWidth(40);

                // T2
                colum = cmdl.getColumn(10);
                colum.setMaxWidth(40);
                colum.setMinWidth(40);
                colum.setWidth(40);

                // T3
                colum = cmdl.getColumn(11);
                colum.setMaxWidth(40);
                colum.setMinWidth(40);
                colum.setWidth(40);

                // T4
                colum = cmdl.getColumn(12);
                colum.setMaxWidth(40);
                colum.setMinWidth(40);
                colum.setWidth(40);

                // T5
                colum = cmdl.getColumn(13);
                colum.setMaxWidth(40);
                colum.setMinWidth(40);
                colum.setWidth(40);

                // T6 @@
                colum = cmdl.getColumn(14);
                colum.setMaxWidth(40);
                colum.setMinWidth(40);
                colum.setWidth(40);

                // PNo
                colum = cmdl.getColumn(15);
                colum.setMaxWidth(32);
                colum.setMinWidth(32);
                colum.setWidth(32);

                // �J�n
                colum = cmdl.getColumn(16);
                colum.setMaxWidth(30);
                colum.setMinWidth(30);
                colum.setWidth(30);
            }
            catch (Throwable e) {
                CZSystem.handleException(e);
            }
        }

        //
        //
        public void valueChanged(ListSelectionEvent e){
            super.valueChanged(e);
        }

        //
        //
        public void setData(int gr,int tbl){

//@@            CZSystem.log("CZPVDataSave","BtConditionTable setData() [" + gr + "][" + tbl + "]");
        }
    }

    /***************************************************************************
     *
     * �a���o�^���ꗗ�F���f��
     *  @@T6�ǉ��ɔ����ύX
     ***************************************************************************/
    public class BtConditionTblMdl extends AbstractTableModel {

        private int TBL_ROW     = 0;
        final   int TBL_COL     = 17;   //@@ 16 -> 17
        private Vector  bt_list     = null;

        final String[] names = {" # "  , "�o�^����" , "�A��" ,  
                                "�i��" , "���c�{"   , "���a" ,
                                "���㒷" , "���d��"   , "�ǎd��" ,
                                "T1" , "T2"   , "T3"  ,
                                "T4" , "T5"   , "T6"  , "PNo" , "�J�n"
                            };

        private Object  data[][];

        BtConditionTblMdl(Vector v){
            super();
            bt_list = v;
            TBL_ROW = bt_list.size();

            data = new Object[TBL_ROW][TBL_COL];

            for(int i = 0 ; i < TBL_ROW ; i++){
                CZSystemBt bt = (CZSystemBt)bt_list.elementAt(i);

                if(null == bt) break;

                data[i][0]  = new Integer(i+1);             //
                data[i][1]  = bt.t_time;                    //�o�^����
                data[i][2]  = new Integer(bt.renban);       //�A��
                data[i][3]  = bt.hinshu;                    //�i��
                data[i][4]  = new Integer(bt.rutubo_kei);   //���c�{�a
                data[i][5]  = new Integer(bt.chokkei);      //���a
                data[i][6]  = new Integer(bt.hikiage_cho);  //���㒷
                data[i][7]  = new Integer(bt.i_sikomi);     //�d����
                data[i][8]  = new Integer(bt.t_sikomi);     //�ǉ��d����
                data[i][9]  = new Integer(bt.no_youkai);    //T1(�n��)
                data[i][10] = new Integer(bt.no_hikiage);   //T2(����)
                data[i][11] = new Integer(bt.no_kaiten);    //T3(��])
                data[i][12] = new Integer(bt.no_toridasi);  //T4(��o)
                data[i][13] = new Integer(bt.no_aturyoku);  //T5(����)
                data[i][14] = new Integer(bt.no_teisu);     //T6(�萔) @@
                data[i][15] = new Integer(bt.pno_start);    //�X�^�[�g�v���Z�X
                data[i][16] = new Integer(bt.p_kaisi);      //�J�n
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
//            data[row][column] = aValue;
        }
    }

    /***************************************************************************
     *
     *       �o�u���ꗗ
     *
     ***************************************************************************/
    class PVTable extends JTable {

        private PVTblMdl model = null;

        PVTable(){
            super();

            try{
                setName("PVTable");
                setAutoCreateColumnsFromModel(true);
                setSelectionMode(ListSelectionModel.MULTIPLE_INTERVAL_SELECTION);

                setLocale(new Locale("ja","JP"));
                setFont(new java.awt.Font("dialog", 0, 12));
                setRowHeight(17);

                model = new PVTblMdl();
                setModel(model);
                DefaultTableColumnModel cmdl = (DefaultTableColumnModel)getColumnModel();
                TableColumn colum = null;

                for(int i=0; i < 128 ; i++){
                    if(prop_PVItemNo[i].equals("1")){
                        addRowSelectionInterval(i,i);
                    }
                }

                // No
                colum = cmdl.getColumn(0);
                colum.setMaxWidth(40);
                colum.setMinWidth(40);
                colum.setWidth(40);

                // ���ږ�
                colum = cmdl.getColumn(1);
                colum.setMaxWidth(80);
                colum.setMinWidth(80);
                colum.setWidth(80);

                // �P��
                colum = cmdl.getColumn(2);
                colum.setMaxWidth(70);
                colum.setMinWidth(70);
                colum.setWidth(70);

                // ����
                colum = cmdl.getColumn(3);
                colum.setMaxWidth(280);
                colum.setMinWidth(280);
                colum.setWidth(280);

CZSystem.log("CZPVDataSave","PVTable Create");
            }
            catch (Throwable e) {
                CZSystem.handleException(e);
            }
        }

        //
        //
        public void valueChanged(ListSelectionEvent e){
//CZSystem.log("CZPVDataSave","PVTable");
            super.valueChanged(e);
        }

        //
        //
        public void setData(int gr,int tbl){

//            CZSystem.log("CZPVDataSave","PVTable setData() [" + gr + "][" + tbl + "]");
        }
    }

    /***************************************************************************
     *
     *       �o�u���ꗗ�F���f��
     *
     ***************************************************************************/
    public class PVTblMdl extends AbstractTableModel {

        private int TBL_ROW     = CZSystemDefine.PV_MAX_LENGTH;
        final   int TBL_COL     = 4;

        final String[] names = {" # " , "���ږ�" , "�P��" , "����" };

        private Object  data[][];

        PVTblMdl(){
            super();

            data = new Object[TBL_ROW][TBL_COL];

            String empty   = new String("");

            for(int i = 0 ; i < TBL_ROW ; i++){
                CZSystemPVName name = CZSystem.getPVName(i);
                if(null == name) break;

                data[i][0] = new Integer(name.k_no);
                data[i][1] = name.k_name;
                data[i][2] = name.unit;
                data[i][3] = name.j_name;
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

    /*
     *
     *       ���l��ݒ肷��TextField
     *
     */
    public class NumeralText extends JTextField {
        private int length = 1;
        private boolean dot = true;

        NumeralText(int len,boolean d){
            super();

            length  = len-1;
            dot = d;

            setFont(new java.awt.Font("dialog", 0, 16));
        }


        //
        //
        //
        public int getValue() {
            int ret = 0;

            String s = getText();
            if(null == s ) return ret;

            try{
                ret = Integer.parseInt(s);
            }
            catch(NumberFormatException e){
                ret = 0;
            }
            return ret;
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

            //
            //
            public void insertString( int offset, String str, AttributeSet a )
                        throws BadLocationException {

                if(length < getLength()) return;

                String validValues = null;

                if(dot){
                    validValues = "0123456789.";
                }
                else {
                    validValues = "0123456789";
                }

                char[] val = str.toCharArray();
                for (int i = 0; i < val.length; i++) {
                    if(validValues.indexOf(val[i]) == -1) return;
                }
                super.insertString( offset, str, a );
            }
        }
    }

    /***************************************************************************
     *
     *       �t�@�C������ݒ肷��TextField
     *
     ***************************************************************************/
    public class FileText extends JTextField {

        FileText(){
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
            String validValues = "0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz.-_";

            //
            //
            public void insertString( int offset, String str, AttributeSet a )
                    throws BadLocationException {

                char[] val = str.toCharArray();
                for (int i = 0; i < val.length; i++) {
                    if(validValues.indexOf(val[i]) == -1) return;
                }
                super.insertString( offset, str, a );
            }
        }
    }

	public class DispBtColorTbl implements Serializable 
	{
	    public String   batch;          //�o�b�`�ԍ�
	    public String	t_name;			//�e�[�u����
	    public int		m_flg;			//�Ԉ����L��
	    public int		m_sumi;			//�Ԉ�����
	    public int		m_sumi_chg;		//�Ԉ����ρi�ύX�l�j
	}
	
}
