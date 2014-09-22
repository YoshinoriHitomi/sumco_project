package cz;

import java.awt.Dimension;
import java.awt.Rectangle;
// add start 2008.10.22
import javax.swing.JScrollBar;
// add end 2008.10.22
import javax.swing.JScrollPane;
import javax.swing.JViewport;

/*
 *  PVグラフ用スクロールパネル
 * @author  (KPK Co.,Ltd.)
 * @version 1.0 (2003/04/01)
 * 2008.10.22 H.Nagamine 初期画面表示状態変更
 */
public class CZPVPanelX extends JScrollPane {

    private CZPVPanelXView view = null;

    CZPVPanelX(int x,int y){
        super();

        try{
            view = new CZPVPanelXView();
            view.setPreferredSize(new Dimension(x,y));

            setName("CZPVPanelX");
            setBounds(100, 620, 600, 40);
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

//@@            CZSystem.log("CZPVPanelX","new");
        }
        catch (Throwable e) {
            CZSystem.handleException(e);
        }
    }
// add start 2008.10.22
    public void setHorizontalScrollBarPosition(int position) {

        JScrollBar jmp_jsb = getHorizontalScrollBar();
        jmp_jsb.setMaximum(250000);
        jmp_jsb.setValue(position);
        setHorizontalScrollBar(jmp_jsb);
    }
// add end 2008.10.22
    //
    //
    //
    public CZPVPanelXView getView(){
        return view;
    }
}

