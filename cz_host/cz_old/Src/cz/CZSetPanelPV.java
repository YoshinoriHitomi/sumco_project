package cz;

import javax.swing.JScrollPane;
import javax.swing.JViewport;

/**********************************************************
 *
 *　　メイン画面：PVグラフ用スクロールパネル
 * @author  (KPK Co.,Ltd.)
 * @version 1.0 (2003/04/01)
 ***********************************************************/
public class CZSetPanelPV extends JScrollPane { 

    private CZSetPanelPVTbl pvTbl = null;

    //_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
    //　コンストラクタ
    //_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
    CZSetPanelPV(){
        super();

        try{
            setName("CZSetPanelPV");
            setVerticalScrollBarPolicy(JScrollPane.VERTICAL_SCROLLBAR_NEVER);
            setHorizontalScrollBarPolicy(JScrollPane.HORIZONTAL_SCROLLBAR_NEVER);
//@@            setBounds(20, 460 , 250, 186);
            setBounds(20, 460 , 250, 190);

           getViewport().setScrollMode(JViewport.BACKINGSTORE_SCROLL_MODE);

           pvTbl = new CZSetPanelPVTbl();
           setViewportView(pvTbl);

//@@           CZSystem.log("CZSetPanelPV","new");
           alterPV();           //@@
        }
        catch (Throwable e) {
            CZSystem.handleException(e);
        }
    }

    //_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
    //　グラフの更新
    //_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
    public void alterPV(){
        pvTbl.alterPV();
    }
}
