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
public class CZPVPanelY2 extends JScrollPane {

    private CZPVPanelY2View view = null;

    CZPVPanelY2(int x,int y){
        super();

        try{

            view = new CZPVPanelY2View();   
            view.setPreferredSize(new Dimension(x,y));          

            setName("CZPVPanelY2");
            setBounds(700, 20, 80, 600);
            setVerticalScrollBarPolicy(JScrollPane.VERTICAL_SCROLLBAR_ALWAYS);
            setHorizontalScrollBarPolicy(JScrollPane.HORIZONTAL_SCROLLBAR_NEVER);

            getViewport().setView(view);
            getViewport().setScrollMode(JViewport.BACKINGSTORE_SCROLL_MODE);

            Rectangle rec = getViewportBorderBounds();
            view.setViewRect(rec);

//@@            CZSystem.log("CZPVPanelY2","new");
        }
        catch (Throwable e) {
            CZSystem.handleException(e);
        }
    }


    //
    //
    //
    public CZPVPanelY2View getView(){
        return view;
    }
}

