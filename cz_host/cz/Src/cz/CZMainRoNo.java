package cz;

import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.util.Vector;

import javax.swing.JComboBox;
    
/***********************************************************
 *   
 *   メイン画面用炉番選択
 * @author  (KPK Co.,Ltd.)
 * @version 1.0 (2003/04/01)
 *   
 ***********************************************************/
public class CZMainRoNo extends JComboBox { 

    // ---------- コンストラクタ ---------------------------
    //
    CZMainRoNo(){
        super();

//@@        CZSystem.log("CZMainRoNo","new"); 
        try{
            setName("JComboBox1");
            setFont(new java.awt.Font("dialog", 0, 24));

            Vector ro = CZSystem.getRoNameList();
            if(null == ro){
                CZSystem.exit(0,"Not Ro No");
            }
            for(int i = 0 ; ro.size() > i ; i++){
				String s = CZSystem.RoKetaChg((String)ro.elementAt(i));	// 20050725 炉：表示桁数変更
                addItem(s);
//                addItem((String)ro.elementAt(i));
            }

            setBounds(20, 20, 100, 40);
            setForeground(java.awt.Color.black);
            setBackground(java.awt.Color.lightGray);
			setFocusable(false);	/* 2007.08.22 */
            addActionListener(new ChgRoNo());
        }
        catch (Throwable e) {
          CZSystem.handleException(e);
        }
    }
    //
    //
    //
    class ChgRoNo implements ActionListener {
        public void actionPerformed(ActionEvent e){
            CZMainRoNo obj = (CZMainRoNo)e.getSource();
//@@            CZSystem.log("CZMainRoNo","ChgRoNo() [" + obj.getSelectedIndex() + "]" );
            CZSystem.chgRo(obj.getSelectedIndex());
        }
    }
}
