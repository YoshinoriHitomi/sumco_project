package cz;

import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.util.Locale;

import javax.swing.JButton;
import javax.swing.JPanel;
import javax.swing.plaf.metal.MetalBorders.Flush3DBorder;

/*******************************************************************************
 *
 *  メイン画面：表示項目用パネル
 * @author  (KPK Co.,Ltd.)
 * @version 1.0 (2003/04/01)
 * Update 2013.10.21 他基地参照機能 (@20131021)
 * Update 2013.10.30 表示切り替え機能 (@20131030)
 ********************************************************************************/
public class CZSetPanel extends JPanel implements Runnable {

    private CZSetPanelSet setTbl    = null;
    private CZSetPanelPV  pvTbl     = null;

    private JButton setTblButton    = null;
    private JButton pvTblButton     = null;
    private JButton chgGraphButton  = null;		// @20131030

    private CZBtSetWin btWin        = null;
    private CZPVSetWin pvWin        = null;

    CZPVPanel _pvPanel;                         // @20131030

    //_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
    // コンストラクタ
    //_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
    CZSetPanel(){   
        super();

        try{
            setName("CZSetPanel");  
            setLayout(null);
            setBorder(new Flush3DBorder());
            setBackground(java.awt.Color.gray);
            setBounds(840, 70, 290, 700);

            setTbl = new CZSetPanelSet();
            add(setTbl, setTbl.getName());  

            pvTbl = new CZSetPanelPV();
            add(pvTbl, pvTbl.getName());

            setTblButton = new JButton("引き上げ条件設定");
            setTblButton.setBounds(20, 420, 250, 30);
            setTblButton.setLocale(new Locale("ja","JP"));  
            setTblButton.setFont(new java.awt.Font("dialog", 0, 18));
            setTblButton.setBorder(new Flush3DBorder());
            setTblButton.setBackground(java.awt.Color.lightGray);
            setTblButton.addActionListener(new SetBtVal());
            add(setTblButton);  

            pvTblButton = new JButton("表示項目設定");
            pvTblButton.setBounds(20, 655, 120, 30);
            pvTblButton.setLocale(new Locale("ja","JP"));
            pvTblButton.setFont(new java.awt.Font("dialog", 0, 18));
            pvTblButton.setBorder(new Flush3DBorder());
            pvTblButton.setBackground(java.awt.Color.lightGray);
            pvTblButton.addActionListener(new SetPVVal());  
            add(pvTblButton);

            // @20131030 表示切り替えボタン
            chgGraphButton = new JButton("表示切り替え");
            chgGraphButton.setBounds(150, 655, 120, 30);
            chgGraphButton.setLocale(new Locale("ja","JP"));
            chgGraphButton.setFont(new java.awt.Font("dialog", 0, 18));
            chgGraphButton.setBorder(new Flush3DBorder());
            chgGraphButton.setBackground(java.awt.Color.lightGray);
            chgGraphButton.addActionListener(new ChgPVVal());  
            add(chgGraphButton);

            btWin = new CZBtSetWin();
            pvWin = new CZPVSetWin();

            CZSystem.log("CZSetPanel","new");
                // @20131021 他基地参照機能
                if(CZSystemDefine.REFERENCE_RUN == CZSystem.getRunLevel()){  // 参照のみの場合、引上げ条件設定画面・表示項目画面は実行しない
                    CZSystem.log("###############################","!!!!!!!!!!!!!!!!!!!");
                    setTblButton.setEnabled(false);
                    pvTblButton.setEnabled(false);
                }  // @20131021

        }
        catch (Throwable e) {
            CZSystem.handleException(e);
        }
    }

    //
    // CZPVPanelクラス Set  @20131030
    //
	public void setPanel(CZPVPanel pvPanel)
	{
		_pvPanel = pvPanel;
	}

    //_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
    // ここから Method
    //_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
    public void run(){  

        CZSystemQueue que = new CZSystemQueue(20);
        CZEventAdapter  adp = new CZEventAdapter(que);  
        CZEventSender.addCZEventListener(adp);  

        while(true){
            try{
                CZEventCL event = (CZEventCL)que.waitObject();  
//@@                CZSystem.log("CZSetPanel run","1");

                if(event.getEvent() == CZEventCL.PV_RECEIVE){
                    setTbl.alterTbl();  
                    pvTbl.alterPV();
                }

                if(event.getEvent() == CZEventCL.RO_CHANGE){
                    setTbl.alterTbl();  
                    pvTbl.alterPV();
                }
            }   
            catch(Exception e){
            }   
//@@            CZSystem.log("CZSetPanel run","2");
        } // while end  
    }

    //_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
    // 引き上げ条件設定画面表示 Class
    //_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
    class SetBtVal implements ActionListener {  
        public void actionPerformed(ActionEvent e){ 
//@@            CZSystem.log("CZSetPanel","SetBtVal");  
            btWin.setDefault();
            btWin.setVisible(true);
        }
    }

    //_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
    // PVグラフ表示項目設定画面表示 Class
    //_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
    class SetPVVal implements ActionListener {  
        public void actionPerformed(ActionEvent e){ 
//@@            CZSystem.log("CZSetPanel","SetPVVal");  
            pvWin.setDefault();
            pvWin.setVisible(true);
        }
    }

    //_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
    // PVグラフ表示項目切替え Class   @20131030
    //_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
    class ChgPVVal implements ActionListener {  
        public void actionPerformed(ActionEvent e){ 
            CZSystem.log("CZSetPanel","ChgPVVal");  
            CZSystem.untenChgView();
            repaint();
			_pvPanel.grpReLoad();
        }
    }
}
