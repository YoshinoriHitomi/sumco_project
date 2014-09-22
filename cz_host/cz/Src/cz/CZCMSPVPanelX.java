package cz;

import java.awt.Dimension;
import java.awt.Rectangle;

import javax.swing.JScrollPane;
import javax.swing.JViewport;

/***********************************************************
 *
 *   集中監視PVグラフ用スクロールパネル
 * @author  (KPK Co.,Ltd.)
 * @version 1.0 (2003/04/01)
 *
 ***********************************************************/
public class CZCMSPVPanelX extends JScrollPane {

    private CZPVPanelXView view = null;

    // ----------  コンストラクタ  -------------------------
    //
    CZCMSPVPanelX(int x,int y){
        super();

        try{
            view = new CZPVPanelXView();
            view.setPreferredSize(new Dimension(x,y));

            setName("CZCMSPVPanelX");
            setBounds(100, 450, 600, 40);
            setVerticalScrollBarPolicy(JScrollPane.VERTICAL_SCROLLBAR_NEVER);
            setHorizontalScrollBarPolicy(JScrollPane.HORIZONTAL_SCROLLBAR_ALWAYS);

            getViewport().setView(view);
            getViewport().setScrollMode(JViewport.BACKINGSTORE_SCROLL_MODE);

            Rectangle rec = getViewportBorderBounds();
            view.setViewRect(rec);

            float w = (float)(rec.width - rec.x) / (float)CZPV.PV_X_SPLIT;
            getHorizontalScrollBar().setBlockIncrement((int)w);

            w = (float)(rec.width - rec.x) / ((float)CZPV.PV_X_SPLIT * 4.0f);
            getHorizontalScrollBar().setUnitIncrement((int)w);

            CZSystem.log("CZCMSPVPanelX","new");
        }
        catch (Throwable e) {
            CZSystem.handleException(e);
        }
    }

    //
    //
    //
    public CZPVPanelXView getView(){
        return view;
    }
}

