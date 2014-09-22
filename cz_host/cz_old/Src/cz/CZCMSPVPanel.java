package cz;

import java.awt.Rectangle;
import java.awt.event.ComponentListener;

import javax.swing.JPanel;
import javax.swing.event.AncestorListener;
import javax.swing.plaf.metal.MetalBorders.Flush3DBorder;

/***********************************************************
 *
 *   集中監視メイン画面ＰＶグラフ表示用パネル
 * @author  (KPK Co.,Ltd.)
 * @version 1.0 (2003/04/01)
 *
 ***********************************************************/
public class CZCMSPVPanel extends JPanel implements AncestorListener,ComponentListener,Runnable {
    
    private CZCMSPVPanelY1      pvY1    = null;
    private CZCMSPVPanelY2      pvY2    = null;
    private CZCMSPVPanelMain    pvMain  = null;
    private CZCMSPVPanelX       pvX     = null;

    //final int X_W     = 5000;
    final int X_W       = 250000;
    final int X_H       = 40;

    final int Y1_W      = 40;
    final int Y1_H      = 2000;

    final int Y2_W      = 80;
    final int Y2_H      = Y1_H;

    final int MAIN_W    = X_W;
    final int MAIN_H    = Y1_H;

    // ---------- コンストラクタ ---------------------------
    //
    CZCMSPVPanel(){
        super();

        try{
            setName("CZPVPanel");
            setLayout(null);
            setBorder(new Flush3DBorder());
            setBackground(java.awt.Color.gray);
            setBounds(20, 320, 800, 500);

            pvY1 = new CZCMSPVPanelY1(Y1_W,Y1_H);
            add(pvY1,pvY1.getName());

            pvMain = new CZCMSPVPanelMain(MAIN_W,MAIN_H);
            add(pvMain,pvMain.getName());

            pvY2 = new CZCMSPVPanelY2(Y2_W,Y2_H);
            add(pvY2,pvY2.getName());
            pvY2.getView().addComponentListener(this);

            pvX  = new CZCMSPVPanelX(X_W,X_H);
            add(pvX,pvX.getName());
            pvX.getView().addComponentListener(this);

            Rectangle rec = pvY2.getViewportBorderBounds();
            pvY2.getView().setLocation(0,rec.height-Y2_H);

            rec = pvY1.getViewportBorderBounds();
            pvY1.getView().setLocation(0,rec.height-Y1_H);

            rec = pvMain.getViewportBorderBounds();
            pvMain.getView().setLocation(0,rec.height-MAIN_H);

            CZSystem.log("CZPVPanel CZPVPanel","new");
        }
        catch (Throwable e) {
            CZSystem.handleException(e);
        }
    }

    //
    //
    //
    public void ancestorMoved(javax.swing.event.AncestorEvent e){

        CZSystem.log("CZPVPanel ancestorMoved","1");
    }
    //
    //
    //
    public void ancestorAdded(javax.swing.event.AncestorEvent e){

        CZSystem.log("CZPVPanel ancestorAdded","1");
    }

    //
    //
    //
    public void ancestorRemoved(javax.swing.event.AncestorEvent e){

        CZSystem.log("CZPVPanel ancestorRemoved","1");
    }


    //
    //
    //
    public void componentMoved(java.awt.event.ComponentEvent e){

        if(pvX.getView() == e.getComponent()){
            CZSystem.log("CZPVPanel componentMoved","1");
            pvX.getView().repaint();
            int x = pvX.getView().getX();
            JPanel view = (JPanel)pvMain.getView();
            view.setLocation(x,view.getY());
            return;
        }

        if(pvY2.getView() == e.getComponent()){
            CZSystem.log("CZCMSPVPanel","PV Move Y");
            pvY2.getView().repaint();
            int y = pvY2.getView().getY();
            JPanel view = (JPanel)pvMain.getView();
            view.setLocation(view.getX(),y);
            view = (JPanel)pvY1.getView();
            view.setLocation(view.getX(),y);
            return;
        }

        CZSystem.log("CZPVPanel componentMoved","2");
    }

    //
    //
    //
    public void componentResized(java.awt.event.ComponentEvent e){

        CZSystem.log("CZPVPanel componentResized","1");
    }

    //
    //
    //
    public void componentShown(java.awt.event.ComponentEvent e){
        CZSystem.log("CZPVPanel componentShown","1");
    }

    //
    //
    //
    public void componentHidden(java.awt.event.ComponentEvent e){
        CZSystem.log("CZPVPanel componentHidden","1");
    }

    //
    //
    //
    public void run(){
    
        CZSystemQueue   que = new CZSystemQueue(20);
        CZEventAdapter  adp = new CZEventAdapter(que);
        CZEventSender.addCZEventListener(adp);

        while(true){
            try{
                CZEventCL event = (CZEventCL)que.waitObject();
                CZSystem.log("CZPVPanel run","1");

                if(event.getEvent() == CZEventCL.PV_RECEIVE){
                    pvMain.getView().setDBData();       //スレッドのタイミングによって
                    //描画されないので追加
                    pvY1.getView().repaint();
                    pvY2.getView().repaint();
                    pvX.getView().repaint();
                    pvMain.getView().repaint();
                }

                if(event.getEvent() == CZEventCL.RO_CHANGE){
                    pvMain.getView().setDBData();
                    pvY1.getView().repaint();
                    pvY2.getView().repaint();
                    pvX.getView().repaint();
                    pvMain.getView().repaint();
                }
            }
            catch(Exception e){

            }

            CZSystem.log("CZPVPanel run","2");
        } // while end
    }
}

