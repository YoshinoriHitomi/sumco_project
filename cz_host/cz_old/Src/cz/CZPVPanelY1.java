package cz;

import java.awt.Dimension;
import java.awt.Rectangle;

import javax.swing.JScrollPane;
import javax.swing.JViewport;

/*
 *  PVグラフ用スクロールパネル
 * @author  (KPK Co.,Ltd.)
 * @version 1.0 (2003/04/01)
 */
public class CZPVPanelY1 extends JScrollPane {

    private CZPVPanelY1View view = null;

    CZPVPanelY1(int x,int y){
        super();

        try{
            view = new CZPVPanelY1View();   
            view.setPreferredSize(new Dimension(x,y));      

            setName("CZPVPanelY1");
            setBounds(20, 20, 80, 600);
            setVerticalScrollBarPolicy(JScrollPane.VERTICAL_SCROLLBAR_NEVER);
            setHorizontalScrollBarPolicy(JScrollPane.HORIZONTAL_SCROLLBAR_NEVER);

            getViewport().setView(view);
            getViewport().setScrollMode(JViewport.BACKINGSTORE_SCROLL_MODE);

            Rectangle rec = getViewportBorderBounds();
            view.setViewRect(rec);

//@@            CZSystem.log("CZPVPanelY1","new");
        }
        catch (Throwable e) {
            CZSystem.handleException(e);
        }

    }

    //
    //
    //
    public CZPVPanelY1View getView(){
        return view;
    }
}

