package cz;

import java.awt.Rectangle;
import java.awt.event.ComponentListener;
// add start 2008.10.22
import javax.swing.JScrollBar;
// add end 2008.10.22
import javax.swing.JPanel;
import javax.swing.event.AncestorListener;
import javax.swing.plaf.metal.MetalBorders.Flush3DBorder;

/*
 *  メイン画面ＰＶグラフ表示用パネル
 * @author  (KPK Co.,Ltd.)
 * @version 1.0 (2003/04/01)
 * 2008.10.22 H.Nagamine 初期画面表示状態変更
 */
public class CZPVPanel extends JPanel implements AncestorListener,ComponentListener,Runnable {
                        
    private CZPVPanelY1     pvY1    = null;
    private CZPVPanelY2     pvY2    = null;
    private CZPVPanelMain   pvMain  = null;
    private CZPVPanelX      pvX     = null;

    final int X_W       = 250000;
    final int X_H       = 40;

    final int Y1_W      = 40;
    final int Y1_H      = 2000;

    final int Y2_W      = 80;
    final int Y2_H      = Y1_H;

    final int MAIN_W    = X_W;
    final int MAIN_H    = Y1_H;

    private boolean   BarMoveFlg;

    CZPVPanel(){
        super();

        try{
            setName("CZPVPanel");
            setLayout(null);
            setBorder(new Flush3DBorder());
            setBackground(java.awt.Color.gray);
//@@            setBounds(20, 70, 800, 680);
            setBounds(20, 70, 800, 670);

            pvY1 = new CZPVPanelY1(Y1_W,Y1_H);
            add(pvY1,pvY1.getName());

            pvMain = new CZPVPanelMain(MAIN_W,MAIN_H);
            add(pvMain,pvMain.getName());

            pvY2 = new CZPVPanelY2(Y2_W,Y2_H);
            add(pvY2,pvY2.getName());
            pvY2.getView().addComponentListener(this);

            pvX  = new CZPVPanelX(X_W,X_H);
            add(pvX,pvX.getName());
            pvX.getView().addComponentListener(this);

            Rectangle rec = pvY2.getViewportBorderBounds();
            pvY2.getView().setLocation(0,rec.height-Y2_H);

            rec = pvY1.getViewportBorderBounds();
            pvY1.getView().setLocation(0,rec.height-Y1_H);

            rec = pvMain.getViewportBorderBounds();
            pvMain.getView().setLocation(0,rec.height-MAIN_H);

            BarMoveFlg = true;

//@@            CZSystem.log("CZPVPanel","new");

        }
        catch (Throwable e) {
          CZSystem.handleException(e);
        }
    }

    //
    //
    //
    public void ancestorMoved(javax.swing.event.AncestorEvent e){

//@@        CZSystem.log("CZPVPanel ancestorMoved","1");
    }
    //
    //
    //
    public void ancestorAdded(javax.swing.event.AncestorEvent e){

//@@        CZSystem.log("CZPVPanel ancestorAdded","1");
    }

    //
    //
    //
    public void ancestorRemoved(javax.swing.event.AncestorEvent e){
                        
//@@        CZSystem.log("CZPVPanel ancestorRemoved","1");
    }

    //
    //
    //
    public void componentMoved(java.awt.event.ComponentEvent e){

        if(pvX.getView() == e.getComponent()){
//@@            CZSystem.log("CZPVPanel componentMoved","1");
            pvX.getView().repaint();
            int x = pvX.getView().getX();
// add start 2008.10.22
            int mx = pvMain.getView().getX();
            CZSystem.log("CZPVPanel componentMoved","main x" + mx);
// add end 2008.10.22
            CZSystem.log("CZPVPanel componentMoved","x" + x);
            CZSystem.log("CZPVPanel componentMoved","BarMoveFlg: " + BarMoveFlg);
            JPanel view = (JPanel)pvMain.getView();
            view.setLocation(x,view.getY());
            BarMoveFlg = false;
            return;
        }

        if(pvY2.getView() == e.getComponent()){
//@@            CZSystem.log("CZPVPanel","componentMoved PV Move Y");
            pvY2.getView().repaint();
            int y = pvY2.getView().getY();
            JPanel view = (JPanel)pvMain.getView();
            view.setLocation(view.getX(),y);
            view = (JPanel)pvY1.getView();
            view.setLocation(view.getX(),y);
        return;
        }
//@@        CZSystem.log("CZPVPanel componentMoved","2");

    }

    //
    //
    //
    public void componentResized(java.awt.event.ComponentEvent e){

//@@        CZSystem.log("CZPVPanel componentResized","1");
    }

    //
    //
    //
    public void componentShown(java.awt.event.ComponentEvent e){
//@@        CZSystem.log("CZPVPanel componentShown","1");
    }

    //
    //
    //
    public void componentHidden(java.awt.event.ComponentEvent e){
//@@        CZSystem.log("CZPVPanel componentHidden","1");
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
//@@                CZSystem.log("CZPVPanel run","1");
                if(event.getEvent() == CZEventCL.PV_RECEIVE){
// chg start 2008.10.22
                    pvMain.getView().setDBData();       //スレッドのタイミングによって
                    if(BarMoveFlg == true){
                        int cnt1 = pvMain.getView().setDBData();
                        int procNo = CZSystem.getProcNo();
                        CZSystem.log("CZPVPanel run ","procNo :" + procNo);
                        if(procNo == 4){
                            if(cnt1 != -1) {
                                CZSystem.log("CZPVPanel run ","cnt1 :" + cnt1);
                                if(cnt1 < (180 * CZPV.getPVGrTimeScale() * 10)){
                                    pvX.setHorizontalScrollBarPosition(0);
                                } else {
                                    CZSystem.log("CZPVPanel run ","CZPV.getPVGrTimeFlag() :" + CZPV.getPVGrTimeFlag());
                                    CZSystem.log("CZPVPanel run ","CZPV.getPVGrTimeScale() :" + CZPV.getPVGrTimeScale());
                                    CZSystem.log("CZPVPanel run ","CZPV.getPVGrLengthScale() :" + CZPV.getPVGrLengthScale());

                                    pvX.setHorizontalScrollBarPosition((cnt1 - (150 * CZPV.getPVGrTimeScale() * 10)) / ( 30 * CZPV.getPVGrTimeScale() * 10 ) * 99 + 50);
                                }
                            } else {
                                pvX.setHorizontalScrollBarPosition(0);
                            }
                         }else{
                            if(cnt1 != -1) {
                                CZSystem.log("CZPVPanel run ","cnt1 :" + cnt1);
                                if(cnt1 < (180 * CZPV.getPVGrTimeScale())){
                                    pvX.setHorizontalScrollBarPosition(0);
                                } else {
                                    CZSystem.log("CZPVPanel run ","CZPV.getPVGrTimeFlag() :" + CZPV.getPVGrTimeFlag());
                                    CZSystem.log("CZPVPanel run ","CZPV.getPVGrTimeScale() :" + CZPV.getPVGrTimeScale());
                                    CZSystem.log("CZPVPanel run ","CZPV.getPVGrLengthScale() :" + CZPV.getPVGrLengthScale());

                                    pvX.setHorizontalScrollBarPosition((cnt1 - (150 * CZPV.getPVGrTimeScale())) / ( 30 * CZPV.getPVGrTimeScale()) * 99 + 50);
                                }
                            } else {
                                pvX.setHorizontalScrollBarPosition(0);
                            }
                        }
                    }
// chg end 2008.10.22
                                        //描画されないので追加
                    pvY1.getView().repaint();
                    pvY2.getView().repaint();
                    pvX.getView().repaint();
                    pvMain.getView().repaint();
                }

                if(event.getEvent() == CZEventCL.RO_CHANGE){
// chg start 2008.10.22
                    pvMain.getView().setDBData();
                    BarMoveFlg = true;
                    if(BarMoveFlg == true){
                        int cnt2 = pvMain.getView().setDBData();
                        int procNo = CZSystem.getProcNo();
                        CZSystem.log("CZPVPanel run ","procNo :" + procNo);
                        if(procNo == 4){
                            if(cnt2 != -1) {
                                CZSystem.log("CZPVPanel run ","cnt2 :" + cnt2);
                                if(cnt2 < (180 * CZPV.getPVGrTimeScale() * 10)){
                                    pvX.setHorizontalScrollBarPosition(0);
                                } else {
                                    CZSystem.log("CZPVPanel run ","CZPV.getPVGrTimeFlag() :" + CZPV.getPVGrTimeFlag());
                                    CZSystem.log("CZPVPanel run ","CZPV.getPVGrTimeScale() :" + CZPV.getPVGrTimeScale());
                                    CZSystem.log("CZPVPanel run ","CZPV.getPVGrLengthScale() :" + CZPV.getPVGrLengthScale());

                                    pvX.setHorizontalScrollBarPosition((cnt2 - (150 * CZPV.getPVGrTimeScale() * 10)) / ( 30 * CZPV.getPVGrTimeScale() * 10 ) * 99 + 50);
                                }
                            } else {
                                pvX.setHorizontalScrollBarPosition(0);
                            }
                         }else{
                            if(cnt2 != -1) {
                                CZSystem.log("CZPVPanel run ","cnt2 :" + cnt2);
                                if(cnt2 < (150 * CZPV.getPVGrTimeScale())){
                                    pvX.setHorizontalScrollBarPosition(0);
                                } else {
                                    CZSystem.log("CZPVPanel run ","CZPV.getPVGrTimeFlag() :" + CZPV.getPVGrTimeFlag());
                                    CZSystem.log("CZPVPanel run ","CZPV.getPVGrTimeScale() :" + CZPV.getPVGrTimeScale());
                                    CZSystem.log("CZPVPanel run ","CZPV.getPVGrLengthScale() :" + CZPV.getPVGrLengthScale());

                                    pvX.setHorizontalScrollBarPosition((cnt2 - (150 * CZPV.getPVGrTimeScale())) / ( 30 * CZPV.getPVGrTimeScale()) * 99 + 50);
                                }
                            } else {
                                pvX.setHorizontalScrollBarPosition(0);
                            }
                        }
                    }
// chg end 2008.10.22
                    pvY1.getView().repaint();
                    pvY2.getView().repaint();
                    pvX.getView().repaint();
                    pvMain.getView().repaint();
                }
            }
            catch(Exception e){

            }
//@@            CZSystem.log("CZPVPanel run","2");
        } // while end

    }

    //
    // グラフ再表示 @20131030
    //
	public void grpReLoad()
	{
        CZSystem.log("CZPVPanel","grpReLoad");
        repaint();
	}

}
