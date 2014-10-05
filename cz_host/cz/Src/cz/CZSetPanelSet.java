package cz;

import javax.swing.JScrollPane;
import javax.swing.JViewport;
import javax.swing.table.JTableHeader;

/*******************************************************************************
 *
 *  メイン画面：引き上げ条件用スクロールパネル
 * @author  (KPK Co.,Ltd.)
 * @version 1.0 (2003/04/01)
 *******************************************************************************/
public class CZSetPanelSet extends JScrollPane {

    private CZSetPanelSetTbl setTbl = null;

    //_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
    // コンストラクタ
    //_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
    CZSetPanelSet(){
        super();

        try{
            setName("CZSetPanelSet");
            setVerticalScrollBarPolicy(JScrollPane.VERTICAL_SCROLLBAR_NEVER);
            setHorizontalScrollBarPolicy(JScrollPane.HORIZONTAL_SCROLLBAR_NEVER);
            setBounds(20, 20, 250, 394);
//@@            setBounds(20, 20, 250, 392);

            getViewport().setScrollMode(JViewport.BACKINGSTORE_SCROLL_MODE);

            setTbl = new CZSetPanelSetTbl();
            JTableHeader tabHead = setTbl.getTableHeader();
            tabHead.setReorderingAllowed(false);
            setViewportView(setTbl);

//@@            CZSystem.log("CZSetPanelSet","new");
            alterTbl();             //@@
        }
        catch (Throwable e) {
            CZSystem.handleException(e);
        }
    }

    //_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
    // テーブルの更新
    //_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
    public void alterTbl(){
        setTbl.alterTbl();
    }
}
