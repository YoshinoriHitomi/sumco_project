package cz;

import java.awt.Dimension;
import java.awt.Rectangle;

import javax.swing.JScrollPane;
import javax.swing.JViewport;

/***********************************************************
 *
 *   �W���Ď�PV�O���t�p�X�N���[���p�l��
 * @author  (KPK Co.,Ltd.)
 * @version 1.0 (2003/04/01)
 *
 ***********************************************************/

public class CZCMSPVPanelY1 extends JScrollPane {

    private CZPVPanelY1View view = null;

    // ----------  �R���X�g���N�^  -------------------------
    //
    CZCMSPVPanelY1(int x,int y){
        super();

        try{
            view = new CZPVPanelY1View();  
            view.setPreferredSize(new Dimension(x,y));   

            setName("CZCMSPVPanelY1");
            setBounds(20, 20, 80, 430);
            setVerticalScrollBarPolicy(JScrollPane.VERTICAL_SCROLLBAR_NEVER);
            setHorizontalScrollBarPolicy(JScrollPane.HORIZONTAL_SCROLLBAR_NEVER);

            getViewport().setView(view);
            getViewport().setScrollMode(JViewport.BACKINGSTORE_SCROLL_MODE);

            Rectangle rec = getViewportBorderBounds();
            view.setViewRect(rec);

            CZSystem.log("CZCMSPVPanelY1","new CZCMSPVPanelY1()");
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
