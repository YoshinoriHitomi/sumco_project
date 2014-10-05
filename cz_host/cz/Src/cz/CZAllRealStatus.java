package cz;

/*
import java.lang.*;
import java.util.*;
import java.text.*;

import java.awt.*;
import java.awt.event.*;

import javax.swing.*;
import javax.swing.text.*;
import javax.swing.table.*;
import javax.swing.plaf.metal.MetalBorders.*;
*/

import java.awt.Color;
import java.awt.Component;
import java.awt.Rectangle;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.text.DecimalFormat;
import java.util.Locale;

import javax.swing.JButton;
import javax.swing.JDialog;
import javax.swing.JLabel;
import javax.swing.JScrollPane;
import javax.swing.JTable;
import javax.swing.JPanel;
import javax.swing.BorderFactory;
import javax.swing.ListSelectionModel;
import javax.swing.plaf.metal.MetalBorders.Flush3DBorder;
import javax.swing.table.AbstractTableModel;
import javax.swing.table.DefaultTableCellRenderer;
import javax.swing.table.DefaultTableColumnModel;
import javax.swing.table.TableColumn;

import czclass.CZNativeMRoState;
import czclass.CZRealNativeWatchItem;

/**
 *   �Ď��󋵕\�����
 * @author  (KPK Co.,Ltd.)
 * @version 1.0 (2003/04/01)
 */
public class CZAllRealStatus extends JDialog {

    private final int   RO_MAX = 99; //48;

	// �ُ� �F
	public static final Color COLOR_ABNORMAL = new Color( 255, 192, 192 );
	// �v���ُ� �F
	public static final Color COLOR_MEASUREABNORMAL = new Color( 192, 255, 255 );
	// ���Ď�
	public static final Color COLOR_NONEWATCH = new Color( 255, 255, 255 );
	// �x�� �F
	public static final Color COLOR_WARN = new Color( 255, 255, 192 );
	// ���� �F
	public static final Color COLOR_NORMAL = new Color( 192, 255, 192 );
	// 4�A���x�� �F
	public static final Color COLOR_4WARN = new Color( 255, 102, 0 );
	// 4�A���ُ� �F
	public static final Color COLOR_4ABNORMAL = new Color( 255, 0, 0 );

    private final int AUTO      = 2;    // 3    ���䃂�[�h

	private final int CHOKKEI  = 24;   // 25    ���a

	private final int SS_TEKAI = 93;   // 94   �V�[�h���x�i�����j
	private final int SR_TEKAI = 94;   // 95   �V�[�h��]�i�����j
	private final int CS_TEKAI = 95;   // 96   ���c�{���x�i�����j
	private final int CR_TEKAI = 96;   // 97   ���c�{��]�i�����j
	private final int HT_TEKAI = 100;  // 101  ���C���q�[�^�P���x�i�����j


    private JLabel      time_lab        = null; //�X�V�����\��
    private JButton     send_button     = null;
    private JButton     cancel_button   = null;

    private JScrollPane st_scpanel      = null; //�X�N���[��
    private StatusTable st_table        = null; //�ғ��󋵃e�[�u��

    private CZNativeMRoState[] status_list;
	/* 2006.07.12 */
    private CZRealNativeWatchItem[] real_list;
	private LegendPanel legendPanel;
	private titlPanel kainyu_panel;
	private titlPanel ias_panel;
	private titlPanel alarm_panel;

    private UpdateThread    updateTh    = null; //�X�V�X���b�h

    /**
    *   �R���X�g���N�^
    */
    CZAllRealStatus(){
        super();

        setTitle("�Ď���");
        setSize(1220+40+10,1000);
        setResizable(false);
        setModal(false);

        getContentPane().setLayout(null);
        // ����n�Q�Ƌ@�\    @20131021
        if(CZSystemDefine.REFERENCE_RUN == CZSystem.getRunLevel()){
            getContentPane().setBackground(CZSystemDefine.DEFAULT_REFERENCE_BACKGROUND_COL);
        }else{
            getContentPane().setBackground(CZSystemDefine.DEFAULT_BACKGROUND_COL);
        }

        JLabel lab = new JLabel("�X�V����",JLabel.CENTER);
        lab.setBounds(10, 20, 100, 24);
        lab.setLocale(new Locale("ja","JP"));
        lab.setFont(new java.awt.Font("dialog", 0, 18));
        lab.setBorder(new Flush3DBorder());
        lab.setForeground(java.awt.Color.black);
        getContentPane().add(lab);

        time_lab = new JLabel("",JLabel.CENTER);
        time_lab.setBounds(110, 20, 200, 24);
        time_lab.setLocale(new Locale("ja","JP"));
        time_lab.setFont(new java.awt.Font("dialog", 0, 18));
        time_lab.setBorder(new Flush3DBorder());
        time_lab.setForeground(java.awt.Color.black);
        getContentPane().add(time_lab);

		// �}��
		legendPanel = new LegendPanel();
		legendPanel.setBounds( 405, 5, 377, 48 );
		getContentPane().add(legendPanel);

		titlPanel kainyu_panel = new titlPanel("�����i�L���j", new Rectangle( 0, 0, 200,18));
		kainyu_panel.setBounds(255, 55, 150, 18);
		getContentPane().add(kainyu_panel);

		titlPanel ias_panel = new titlPanel("IAS(In-situ Alert System)", new Rectangle( 0, 0, 480, 18));
		ias_panel.setBounds(405, 55, 480, 18);                                               
		getContentPane().add(ias_panel);

		titlPanel alarm_panel = new titlPanel("�A���[��", new Rectangle( 0, 0, 350, 18));
		alarm_panel.setBounds(885, 55, 350, 18);                                               
		getContentPane().add(alarm_panel);

        //
        st_table = new StatusTable();
        st_scpanel = new JScrollPane(st_table);
        st_scpanel.setBounds(10, 50+18+5, 1180+40+20, 837+34);
        st_scpanel.setBorder(new Flush3DBorder());
        getContentPane().add(st_scpanel);

        updateTh = new UpdateThread();
        updateTh.setPriority(Thread.MIN_PRIORITY);
        updateTh.start();
        
        CZSystem.log("CZAllRealStatus","�Ď��󋵁@�\��");
    }


    //
    //
    //
    public boolean setDefault(){

        return true;
    }

    //
    //
    //
    private boolean updateStatus(){

		/* �V�X�e����Ԏ擾 */
        status_list = null;
        status_list = CZSystem.CZNativeMRoStateGet();

        if(null == status_list) return false;

        int size = status_list.length;
        if(1 > size) return false;
CZSystem.log("CZAllRealStatus","Size[" + size + "]");

        st_table.setData();

        String tm = CZSystem.getDateTime();
        time_lab.setText(tm);

        return true;
    } 


    /**
    *
    */
    class SendButton implements ActionListener {
        public void actionPerformed(ActionEvent ev){
            updateStatus();
            CZSystem.sleep(5000);
        }
    }


    /**
    *
    */
    class CancelButton implements ActionListener {
        public void actionPerformed(ActionEvent ev){
            setDefault();
            setVisible(false);
        }
    }


    /**
    *   �X�V�X���b�h
    */
    class UpdateThread extends Thread {

        UpdateThread(){

        }

        public void run(){
//@@            CZSystem.log("CZAllRealStatus","UpdateThread START");

            while(true){
                updateStatus();
                CZSystem.sleep(1000 * 10);	/* 10s */
            } // while end
        }
    }


    /**
    *   �󋵕\���p�e�[�u��
    */
    public class StatusTable extends JTable {

        private StatusModel model = null;
            
        StatusTable(){
            super();

            setName("StatusTable");
            setAutoCreateColumnsFromModel(true);
            setSelectionMode(ListSelectionModel.SINGLE_SELECTION);
            setLocale(new Locale("ja","JP"));
            setFont(new java.awt.Font("dialog", 0, 16));
            setRowHeight(17);
            setRowSelectionAllowed(false);

            model = new StatusModel();
            setModel(model);

            DefaultTableColumnModel cmdl = (DefaultTableColumnModel)getColumnModel();
            TableColumn     colum = null;
            ColorRender ren   = null;

            //�F
            ren = new ColorRender();
            ren.setHorizontalAlignment(ren.CENTER);
            colum = cmdl.getColumn(0);
            colum.setMaxWidth(40+10);
            colum.setMinWidth(40+10);
            colum.setWidth(40+10);
            colum.setCellRenderer(ren);

            //�v���Z�X
            ren = new ColorRender();
            colum = cmdl.getColumn(1);
            colum.setMaxWidth(90);
            colum.setMinWidth(90);
            colum.setWidth(90);
            colum.setCellRenderer(ren);

            //���[�h
            ren = new ColorRender();
            ren.setHorizontalAlignment(ren.CENTER);
            colum = cmdl.getColumn(2);
            colum.setMaxWidth(50);
            colum.setMinWidth(50);
            colum.setWidth(50);
            colum.setCellRenderer(ren);

            //���a
            ren = new ColorRender();
            ren.setHorizontalAlignment(ren.CENTER);
            colum = cmdl.getColumn(3);
            colum.setMaxWidth(55);
            colum.setMinWidth(55);
            colum.setWidth(55);
            colum.setCellRenderer(ren);

            //����(S.S)
            ren = new ColorRender();
            ren.setHorizontalAlignment(ren.CENTER);
            colum = cmdl.getColumn(4);
            colum.setMaxWidth(30);
            colum.setMinWidth(30);
            colum.setWidth(30);
            colum.setCellRenderer(ren);

            //����(S.R)
            ren = new ColorRender();
            ren.setHorizontalAlignment(ren.CENTER);
            colum = cmdl.getColumn(5);
            colum.setMaxWidth(30);
            colum.setMinWidth(30);
            colum.setWidth(30);
            colum.setCellRenderer(ren);

            //����(C.S)
            ren = new ColorRender();
            ren.setHorizontalAlignment(ren.CENTER);
            colum = cmdl.getColumn(6);
            colum.setMaxWidth(30);
            colum.setMinWidth(30);
            colum.setWidth(30);
            colum.setCellRenderer(ren);

            //����(C.R)
            ren = new ColorRender();
            ren.setHorizontalAlignment(ren.CENTER);
            colum = cmdl.getColumn(7);
            colum.setMaxWidth(30);
            colum.setMinWidth(30);
            colum.setWidth(30);
            colum.setCellRenderer(ren);

            //����(H.T)
            ren = new ColorRender();
            ren.setHorizontalAlignment(ren.CENTER);
            colum = cmdl.getColumn(8);
            colum.setMaxWidth(30);
            colum.setMinWidth(30);
            colum.setWidth(30);
            colum.setCellRenderer(ren);

            //IAS-1
            ren = new ColorRender();
            ren.setHorizontalAlignment(ren.LEFT);
            colum = cmdl.getColumn(9);
            colum.setMaxWidth(160);
            colum.setMinWidth(160);
            colum.setWidth(160);
            colum.setCellRenderer(ren);

            //IAS-2
            ren = new ColorRender();
            ren.setHorizontalAlignment(ren.LEFT);
            colum = cmdl.getColumn(10);
            colum.setMaxWidth(160);
            colum.setMinWidth(160);
            colum.setWidth(160);
            colum.setCellRenderer(ren);

            //IAS-3
            ren = new ColorRender();
            ren.setHorizontalAlignment(ren.LEFT);
            colum = cmdl.getColumn(11);
            colum.setMaxWidth(160);
            colum.setMinWidth(160);
            colum.setWidth(160);
            colum.setCellRenderer(ren);

            //�A���[��No
            ren = new ColorRender();
            ren.setHorizontalAlignment(ren.RIGHT);
            colum = cmdl.getColumn(12);
            colum.setMaxWidth(55);
            colum.setMinWidth(55);
            colum.setWidth(55);
            colum.setCellRenderer(ren);

            //�A���[��
            ren = new ColorRender();
            ren.setHorizontalAlignment(ren.LEFT);
            colum = cmdl.getColumn(13);
            colum.setMaxWidth(250);
            colum.setMinWidth(250);
            colum.setWidth(250);
            colum.setCellRenderer(ren);

            //�A���[�����
            ren = new ColorRender();
            ren.setHorizontalAlignment(ren.CENTER);
            colum = cmdl.getColumn(14);
            colum.setMaxWidth(40);
            colum.setMinWidth(40);
            colum.setWidth(40);
            colum.setCellRenderer(ren);
        }

        //
        //
        //
        public void setData(){

			int iSetCnt;
			int iLp;

            if(null == status_list) return;
            int size = status_list.length;


            if(1 > size) return;

//			if ( size > RO_MAX) size = RO_MAX;

            float[] pv;
            DecimalFormat f1 = new DecimalFormat("0");
            DecimalFormat f2 = new DecimalFormat("###0.0");
            DecimalFormat f3 = new DecimalFormat("#0.0000");

            ColorRender cell;

            for(int i = 0 ; i < size ; i++){
                CZNativeMRoState st = status_list[i];

                //PV�f�[�^
                pv = st.getData();

                //�F��
				String s = CZSystem.RoKetaChg(st.getRoName());	// 20050725 �F�F�\�������ύX
                model.setValueAt(s, i,0);

				/* 2006.07.12 */
				real_list = CZSystem.CZNativeRealStateGet(st.getRoName());


                //�v���Z�X
                model.setValueAt(new ColorString(CZSystem.getProcName(st.getP_no()),java.awt.Color.blue), i,1);
                if(st.getDown()) model.setValueAt(new ColorString("DOWN",java.awt.Color.red), i,1);

                //���[�h
                cell = (ColorRender)getCellRenderer(i,2);
                int mode = (int)pv[AUTO];
                switch(mode){
                    case CZSystemDefine.PROC_MANUAL :
                        model.setValueAt(new ColorString(
                        CZSystemDefine.PROC_MODE[CZSystemDefine.PROC_MANUAL],java.awt.Color.red),i,2);
                        break;

                    case CZSystemDefine.PROC_AUTO :
                        model.setValueAt(new ColorString(
                        CZSystemDefine.PROC_MODE[CZSystemDefine.PROC_AUTO],java.awt.Color.blue),i,2);
                        break;

                    default : 
                    model.setValueAt(new ColorString("�s  ��",java.awt.Color.black),i,2);
                        break;
                }

                //���a
                model.setValueAt(f2.format(pv[CHOKKEI]),i,3);

                //����
				if (pv[SS_TEKAI] != (float)0.0) {
                    model.setValueAt(new ColorString("",java.awt.Color.blue,java.awt.Color.blue),i,4);
                } else {
                    model.setValueAt(new ColorString(""),i,4);
                }

				if (pv[SR_TEKAI] != (float)0.0) {
                    model.setValueAt(new ColorString("",java.awt.Color.blue,java.awt.Color.blue),i,5);
                } else {
                    model.setValueAt(new ColorString(""),i,5);
                }

				if (pv[CS_TEKAI] != (float)0.0) {
                    model.setValueAt(new ColorString("",java.awt.Color.blue,java.awt.Color.blue),i,6);
                } else {
                    model.setValueAt(new ColorString(""),i,6);
                }

				if (pv[CR_TEKAI] != (float)0.0) {
                    model.setValueAt(new ColorString("",java.awt.Color.blue,java.awt.Color.blue),i,7);
                } else {
                    model.setValueAt(new ColorString(""),i,7);
                }

				if (pv[HT_TEKAI] != (float)0.0) {
                    model.setValueAt(new ColorString("",java.awt.Color.blue,java.awt.Color.blue),i,8);
                } else {
                    model.setValueAt(new ColorString(""),i,8);
                }

				for (iLp = 9; iLp < 12; iLp++) {
	                 model.setValueAt(new ColorString(""),i,iLp);
				}

                //�Ď���
				if (real_list != null) {
					if (real_list.length != 0) {
						iSetCnt = 9;
						for (iLp=0; (iLp < real_list.length) && (iSetCnt <= 11); iLp++) {
							int iNowState = real_list[iLp].getNowState();
							switch ( iNowState )
							{
								case 1:		/*�x��*/
                    model.setValueAt(new ColorString(real_list[iLp].getItemName(),java.awt.Color.black,COLOR_WARN),i,iSetCnt);
									iSetCnt++;
									break;
								case 2:		/*�ُ�*/
                    model.setValueAt(new ColorString(real_list[iLp].getItemName(),java.awt.Color.black,COLOR_ABNORMAL),i,iSetCnt);
									iSetCnt++;
									break;
/*-------------------------------------------
								case 5:		/#4�A���x��#/
                    model.setValueAt(new ColorString(real_list[iLp].getItemName(),java.awt.Color.black,COLOR_4WARN),i,iSetCnt);
									iSetCnt++;
									break;
  -------------------------------------------*/
								case 6:		/*4�A���ُ�*/
                    model.setValueAt(new ColorString(real_list[iLp].getItemName(),java.awt.Color.black,COLOR_4ABNORMAL),i,iSetCnt);
									iSetCnt++;
									break;
							}
						}
					}
				}

				//�A���[��
//CZSystem.log("test", "alam[" + st.getAlm_no() + "][" + st.getAlm_msg() + "][" + st.getAlm_state() + "]");
				if ( st.getAlm_no() != 0)
	            {
				    model.setValueAt(f1.format(st.getAlm_no()),i,12);
				}
				else
				    model.setValueAt("",i,12);

                model.setValueAt(new ColorString(st.getAlm_msg(),java.awt.Color.black),i,13);

				if ( st.getAlm_state().equals("����") == true)
				    model.setValueAt(new ColorString(st.getAlm_state(),java.awt.Color.red),i,14);
				else
				    model.setValueAt(new ColorString(st.getAlm_state(),java.awt.Color.blue),i,14);

            } // for end
            repaint();
        } 


        /**
        *   �󋵕\���e�[�u���̃��f��
        */
        public class StatusModel extends AbstractTableModel {
            final   int TBL_ROW = RO_MAX;
            final   int TBL_COL = 15;

            final   String[] names = {
                        "�F" , "Proc", "Mode", "���a", "S.S","S.R",
						"C.S","C.R","H.T","�Ď����ږ�-1","�Ď����ږ�-2","�Ď����ږ�-3","E-Code","���b�Z�[�W","���"
            };

            private Object data[][];

            StatusModel(){
                super();

                data = new Object[TBL_ROW][TBL_COL];

                for(int i = 0 ; i < TBL_ROW ; i++){
	                for(int i2 = 0 ; i2 < TBL_COL ; i2++){
	                    data[i][i2]  = "";
	                }
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
        } // public class StatusModel extends AbstractTableModel


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

                String s = "";

                if(String.class == value.getClass()) s = (String)value; 

                if(ColorString.class == value.getClass()){
                    ColorString cl = (ColorString)value;
                    s = cl.getText();
                    setForeground(cl.getColor());
					setBackground(cl.getBkColor());
                }
                super.getTableCellRendererComponent(table,
                                                    s,
                                                    isSelected,
                                                    hasFocus,
                                                    row,column);
                return(this);
            }
        } //class ColorRender extends DefaultTableCellRenderer

        /**
        *
        */
        public class ColorString {
            Color color = java.awt.Color.black;
            Color bkColor = COLOR_NONEWATCH;
            String string = "";

            ColorString(String s){
                string = s;
            }

            ColorString(String s,Color c){
                string = s;
                color = c;
				bkColor = COLOR_NONEWATCH;
            }

            ColorString(String s,Color c,Color c2){
                string = s;
                color = c;
				bkColor = c2;
            }

            public String getText(){
                return string;
            }

            public String toString(){
                return string;
            }

            public Color getColor(){
                return color;
            }

            public Color getBkColor(){
                return bkColor;
            }
        } //public class ColorString
    } //public class StatusTable extends JTable

	/**
	* �}��\���p�l���N���X
	* @return LegendPanel
	*
	*/
	public class LegendPanel extends JPanel {
		LegendPanel(){
			super();
			setLayout(null);
			setBorder( BorderFactory.createTitledBorder(new Flush3DBorder(),"�}��",1,0,new java.awt.Font("dialog", 0, 12),java.awt.Color.black) );

			// ����n�Q�Ƌ@�\    @20131021
			if(CZSystemDefine.REFERENCE_RUN == CZSystem.getRunLevel()){
				setBackground(CZSystemDefine.DEFAULT_REFERENCE_BACKGROUND_COL);
			}else{
				setBackground(CZSystemDefine.DEFAULT_BACKGROUND_COL);
			}

			add( createLegendLabel( new Rectangle( 64, 24, 33, 17), "�x��" ) );
			add( createLegendPanel( new Rectangle( 16, 24, 41, 17 ), COLOR_WARN ) );
			
			add( createLegendLabel( new Rectangle( 176, 24, 33, 17 ), "�ُ�" ) );
			add( createLegendPanel( new Rectangle( 128, 24, 41, 17 ), COLOR_ABNORMAL ) );

			add( createLegendLabel( new Rectangle( 288, 24, 60, 17 ), "�S�A�ُ�" ) );
			add( createLegendPanel( new Rectangle( 240, 24, 41, 17 ), COLOR_4ABNORMAL ) );
			
/*			add( createLegendLabel( new Rectangle( 296, 24, 49, 17 ), "���Ď�" ) );
			add( createLegendPanel( new Rectangle( 248, 24, 41, 17 ), CommonGui.COLOR_NONEWATCH ) );
			add( createLegendLabel( new Rectangle( 176, 48, 60, 17 ), "����" ) );
			add( createLegendPanel( new Rectangle( 128, 48, 41, 17 ), CommonGui.COLOR_NORMAL ) );
*/
		}
		
		private JLabel createLegendLabel( Rectangle rect, String title ) {
			JLabel label = new JLabel( title );
			label.setBounds( rect );
	        label.setLocale(new Locale("ja","JP"));
	        label.setFont(new java.awt.Font("dialog", 0, 12));
			label.setForeground( java.awt.Color.black );
			
			
			return label;
		}
		
		private JPanel createLegendPanel( Rectangle rect, Color color ) {
			JPanel panel = new JPanel();
			panel.setBounds( rect );
			panel.setLocale(new Locale("ja","JP"));
			panel.setBorder(BorderFactory.createLineBorder(java.awt.Color.black));
			panel.setBackground( color );
			
			return panel;
		}
		
	}	// LegendPanel

	/**
	* �^�C�g���\���p�l���N���X
	* @return LegendPanel
	*
	*/
	public class titlPanel extends JPanel {
		titlPanel(String title, Rectangle rect){
			super();
			setLayout(null);
			setBorder( new Flush3DBorder() );
			setBackground(java.awt.Color.lightGray);

			add( createLegendLabel( rect , title ) );
		}
		
		private JLabel createLegendLabel( Rectangle rect, String title ) {
			JLabel label = new JLabel( title, JLabel.CENTER );
			label.setBounds( rect );
	        label.setLocale(new Locale("ja","JP"));
	        label.setFont(new java.awt.Font("dialog", 0, 12));
			label.setForeground( java.awt.Color.black );
			return label;
		}
	}	// titlePanel
}
