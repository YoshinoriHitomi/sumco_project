package cz;

import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.util.Locale;

import javax.swing.JButton;
import javax.swing.JPanel;
import javax.swing.plaf.metal.MetalBorders.Flush3DBorder;

/*******************************************************************************
 *
 *   表示項目用パネル
 * @author  (KPK Co.,Ltd.)
 * @version 1.0 (2003/04/01)
 *
 *******************************************************************************/
public class CZCMSSetPanel extends JPanel implements Runnable {

    private CZSetPanelSet setTbl = null;
    private CZSetPanelPV  pvTbl  = null;

    private JButton setTblButton = null;
    private JButton pvTblButton  = null;

    private CZBtSetWin btWin     = null;
    private CZPVSetWin pvWin     = null;

    //_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
    // コンストラクタ
    //_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
    CZCMSSetPanel(){
        super();

        try{
            setName("CZCMSSetPanel");
            setLayout(null);
            setBorder(new Flush3DBorder());
            setBackground(java.awt.Color.gray);
            setBounds(840, 88, 290, 728);

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
            pvTblButton.setBounds(20, 655, 250, 30);
            pvTblButton.setLocale(new Locale("ja","JP"));
            pvTblButton.setFont(new java.awt.Font("dialog", 0, 18));
            pvTblButton.setBorder(new Flush3DBorder());
            pvTblButton.setBackground(java.awt.Color.lightGray);
            pvTblButton.addActionListener(new SetPVVal());
            add(pvTblButton);

            btWin = new CZBtSetWin();
            pvWin = new CZPVSetWin();

            CZSystem.log("CZCMSSetPanel","new");
        }
        catch (Throwable e) {
            CZSystem.handleException(e);
        }
    }

    //_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
    // method
    //_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
    public void run(){
        CZSystemQueue   que = new CZSystemQueue(20);
        CZEventAdapter  adp = new CZEventAdapter(que);
        CZEventSender.addCZEventListener(adp);

        while(true){
            try{
                CZEventCL event = (CZEventCL)que.waitObject();
                CZSystem.log("CZCMSSetPanel run","1");
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
            CZSystem.log("CZCMSSetPanel run","2");
        } // while end
    }

    //_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
    // 引き上げ条件設定画面表示
    //_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
    class SetBtVal implements ActionListener {

        public void actionPerformed(ActionEvent e){
            CZSystem.log("CZCMSSetPanel","SetBtVal");
            btWin.setDefault();
            btWin.setVisible(true);
        }
    }

    //_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
    // PVグラフ表示項目設定画面表示
    //_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
    class SetPVVal implements ActionListener {

        public void actionPerformed(ActionEvent e){
            CZSystem.log("CZCMSSetPanel","SetPVVal");
            pvWin.setDefault();
            pvWin.setVisible(true);
        }
    }
}
