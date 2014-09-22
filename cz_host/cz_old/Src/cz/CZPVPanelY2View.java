package cz;

import java.awt.Dimension;
import java.awt.Graphics;
import java.awt.Rectangle;

import javax.swing.JPanel;

/*
 *   メイン画面ＰＶグラフ表示用パネル
 * @author  (KPK Co.,Ltd.)
 * @version 1.0 (2003/04/01)
 */
public class CZPVPanelY2View extends JPanel {
    private Rectangle rec = null;

    private boolean untenflg = true;	// @20131030

    CZPVPanelY2View(){
        super();

        try{
            setName("CZPVPanelY2View");
            setLayout(null);
            setBackground(CZPV.PV_BACK_COLOR);

//@@            CZSystem.log("CZPVPanelY2View","new");

        }
        catch (Throwable e) {
            CZSystem.handleException(e);
        }

    }

    //
    //
    //
    public void setViewRect(Rectangle r){
        rec = r;
    }

    //
    //
    //
    public void paint(Graphics g){

        Dimension d = getSize(null);

        g.setColor(CZPV.PV_BACK_COLOR);
        g.fillRect(0,0,d.width,d.height);

        float leng[]  = new float[2];
        float inc[]   = new float[2];
        int   y_pos[] = new int[2];

// @20131030
        untenflg = CZSystem.untenView();
        if (untenflg == true){
	        for(int i = 0 ; i < 2 ; i++){
	            inc[i]  = (CZPV.getPVGrMax(i+3) - CZPV.getPVGrMin(i+3)) / CZPV.PV_Y_SPLIT;
	            leng[i] = CZPV.getPVGrMin(i+3);
	        }
        } else {
	        for(int i = 5 ; i < 7 ; i++){
	            inc[i-5]  = (CZPV.getPVGrMax(i+3) - CZPV.getPVGrMin(i+3)) / CZPV.PV_Y_SPLIT;
	            leng[i-5] = CZPV.getPVGrMin(i+3);
	        }
        }
// @20131030

        y_pos[0] = -20;
        y_pos[1] = -5;

        float  h = (float)(rec.height - rec.y) / (float)CZPV.PV_Y_SPLIT;
        for(float y = (float)d.height ; 0 < y ; y-=h){
            g.setColor(CZPV.PV_MEM_COLOR);
            g.drawLine(0,(int)y,d.width,(int)y);

// @20131030
//CZSystem.log("CZPVPanelY2View paint","メイン画面ＰＶグラフ表示用パネル    Ｙ軸右 表示２");
            if (untenflg == true){
	            for(int i = 0 ; i < 2 ; i++){
	                g.setColor(CZPV.PV_COLOR[i+3]);
	                g.drawString(new String(leng[i] + " " + CZPV.getPVGrUnit(i+3)),10,(int)y+y_pos[i]);
	                leng[i]+=inc[i];
	            }
            } else {
	            for(int i = 5 ; i < 7 ; i++){
	                g.setColor(CZPV.PV_COLOR2[i+3]);
	                g.drawString(new String(leng[i-5] + " " + CZPV.getPVGrUnit(i+3)),10,(int)y+y_pos[i-5]);
	                leng[i-5]+=inc[i-5];
	            }
            }
// @20131030

        }
    }
}
