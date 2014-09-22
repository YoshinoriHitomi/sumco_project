package cz;

import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.text.DecimalFormat;
import java.util.Locale;
import java.util.Vector;

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
 *　　メイン画面：炉選択画面
 * @author  (KPK Co.,Ltd.)
 * @version 1.0 (2003/04/01)
 * 2008.09.16 H.Nagamine 炉番号選択方法変更
 ***********************************************************/
public class CZRoSelectWin extends JDialog {
    private JRadioButton ro_sel_Pad[]     = new JRadioButton[100];
    private JLabel       ro_name_lab[]    = new JLabel[100];
    private JButton      send_button      = null;
    private JButton      cancel_button    = null;
    private JLabel       select_ro_lab    = null;
    private int          ro_index         = 0;

    private int          xPos             = 20;
    private int          yPos             = 80;

    //_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
    //　ここからコンストラクタ
    //_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
    CZRoSelectWin(int X,int Y){
        super();

        setTitle("炉選択");
//        setSize(1030,560);
//        setLocation(20,130);
        setLocation(X+25,Y+130);
        setResizable(false);
        setModal(true);

        getContentPane().setLayout(null);
        // 他基地参照機能    @20131021
        if(CZSystemDefine.REFERENCE_RUN == CZSystem.getRunLevel()){
            getContentPane().setBackground(CZSystemDefine.DEFAULT_REFERENCE_BACKGROUND_COL);
        }else{
            getContentPane().setBackground(CZSystemDefine.DEFAULT_BACKGROUND_COL);
        }

		Vector ro = CZSystem.getRoNameList();
		
// chg start 2008.09.16
//		setSize(100*((ro.size()+1)/10)+30,560);
		setSize(100*((ro.size()+1)/10)+30,530);
// chg end 2008.09.16
		
		ButtonGroup robutton_group = new ButtonGroup();
		
		for(int i = 0; i < ro.size(); i++){
			String s = CZSystem.RoKetaChg((String)ro.elementAt(i));
			ro_sel_Pad[i] = new JRadioButton(s);
			ro_sel_Pad[i].setBorder(new Flush3DBorder());
			ro_sel_Pad[i].setBounds(xPos + (i/10)*100, yPos + (i*40) - (i/10)*400, 80, 30);
			ro_sel_Pad[i].setFont(new java.awt.Font("dialog", 0, 18));
			ro_sel_Pad[i].addActionListener(new SelRoNo());
			getContentPane().add(ro_sel_Pad[i]);
			robutton_group.add(ro_sel_Pad[i]);
		}
		
		String roN = CZMain.roName_lab.getText();
		
		int rind = CZSystem.getRoIndex(roN);
		
		CZSystem.log("CZRoSelectWin","Ro INDEX :" + rind);
		
		ro_sel_Pad[rind].setSelected(true);
		
// del start 2008.09.16
//		send_button = new JButton("実  行");
////		send_button.setBounds(410, 480, 100, 30);
//		send_button.setBounds((100*((ro.size()+1)/10)+30)/2-100, 480, 100, 30);
//		send_button.setLocale(new Locale("ja","JP"));
//		send_button.setFont(new java.awt.Font("dialog", 1, 22));
//		send_button.setBorder(new Flush3DBorder());
//		send_button.setBackground(java.awt.Color.lightGray);
//		send_button.addActionListener(new ChgRoName());  
//		getContentPane().add(send_button);
// del end 2008.09.16

// del start 2008.09.16
//		cancel_button = new JButton("キャンセル");
////		cancel_button.setBounds(510, 480, 100, 30);
//		cancel_button.setBounds((100*((ro.size()+1)/10)+30)/2, 480, 100, 30);
//		cancel_button.setLocale(new Locale("ja","JP"));
//		cancel_button.setFont(new java.awt.Font("dialog", 1, 18));
//		cancel_button.setBorder(new Flush3DBorder());
//		cancel_button.setBackground(java.awt.Color.lightGray);
//		cancel_button.addActionListener(new Cancel());  
//		getContentPane().add(cancel_button);
// del end 2008.09.16

		select_ro_lab = new JLabel(ro_sel_Pad[rind].getText(),JLabel.CENTER);
//		select_ro_lab.setBounds(460, 20, 100, 50);
		select_ro_lab.setBounds((100*((ro.size()+1)/10)+30)/2-50, 20, 100, 50);
		select_ro_lab.setLocale(new Locale("ja","JP"));
		select_ro_lab.setFont(new java.awt.Font("dialog", 1, 28));
		select_ro_lab.setBorder(new Flush3DBorder());
		select_ro_lab.setForeground(java.awt.Color.black);
		getContentPane().add(select_ro_lab);

	}

	class SelRoNo implements ActionListener {
		public void actionPerformed(ActionEvent e){
		
			for(int rec = 0; rec < 100; rec++){
				if(true == ro_sel_Pad[rec].isSelected()){
					CZSystem.log("CZRoSelectWin","Ro Name :" + ro_sel_Pad[rec].getText());
					select_ro_lab.setText(ro_sel_Pad[rec].getText());
					ro_index = rec;
// chg start 2008.09.16
//					return;
					break;
// chg end 2008.09.16
				}
			}
// add start 2008.09.16
			CZSystem.chgRo(ro_index);
			CZMain.roName_lab.setText(ro_sel_Pad[ro_index].getText());
			setVisible(false);
// add end 2008.09.16

		}
	}
	
// del start 2008.09.16
//	class ChgRoName implements ActionListener {
//		public void actionPerformed(ActionEvent e){
//			
//			CZSystem.chgRo(ro_index);
//			CZMain.roName_lab.setText(ro_sel_Pad[ro_index].getText());
//			setVisible(false);
//			
//		}
//	}
// del end 2008.09.16

// del start 2008.09.16	
//	class Cancel implements ActionListener {
//		public void actionPerformed(ActionEvent e){
//			
//			setVisible(false);
//			
//		}
//	}
// del end 2008.09.16
}
