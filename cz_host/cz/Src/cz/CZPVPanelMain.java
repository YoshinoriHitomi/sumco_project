package cz;

import java.awt.Dimension;
import java.awt.Rectangle;

import javax.swing.JScrollPane;
import javax.swing.JViewport;

/**
 *  PVグラフ用スクロールパネル
 * @author  (KPK Co.,Ltd.)
 * @version 1.0 (2003/04/01)
 */
public class CZPVPanelMain extends JScrollPane {

    private CZPVPanelMainView view = null;

    CZPVPanelMain(int x,int y){
        super();

        try{

            view = new CZPVPanelMainView();
            view.setPreferredSize(new Dimension(x,y));


            setName("CZPVPanelMain");
            setBounds(100, 20, 600, 600);
            setVerticalScrollBarPolicy(JScrollPane.VERTICAL_SCROLLBAR_NEVER);
            setHorizontalScrollBarPolicy(JScrollPane.HORIZONTAL_SCROLLBAR_NEVER);

            getViewport().setView(view);
            getViewport().setScrollMode(JViewport.BACKINGSTORE_SCROLL_MODE);

            Rectangle rec = getViewportBorderBounds();
            view.setViewRect(rec);

//@@            CZSystem.log("CZPVPanelMain","new");
        }
        catch (Throwable e) {
            CZSystem.handleException(e);
        }

    }


    //
    //
    //
    public CZPVPanelMainView getView(){
        return view;
    }
}

