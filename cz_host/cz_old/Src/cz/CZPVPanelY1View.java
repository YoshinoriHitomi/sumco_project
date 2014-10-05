package cz;

import java.awt.Dimension;
import java.awt.Graphics;
import java.awt.Rectangle;

import javax.swing.JPanel;

/*
 *   メイン画面ＰＶグラフ表示用パネル    Ｙ軸左
 * @author  (KPK Co.,Ltd.)
 * @version 1.0 (2003/04/01)
 * Update 2013/10/30 表示切り替え機能 (@20131030)
 */
public class CZPVPanelY1View extends JPanel {
    private Rectangle rec = null;

    private boolean untenflg = true;	// @20131030

    CZPVPanelY1View(){
        super();

        try{
            setName("CZPVPanelY1View");
            setLayout(null);
            setBackground(CZPV.PV_BACK_COLOR);

//@@            CZSystem.log("CZPVPanelY1View","new");
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

        float leng[]  = new float[3];
        float inc[]   = new float[3]; 
        int   y_pos[] = new int[3];

// @20131030
        untenflg = CZSystem.untenView();
CZSystem.log("CZPVPanelY1View paint"," untenFlg :" + untenflg);
        if (untenflg == true){
//CZSystem.log("CZPVPanelY1View paint","メイン画面ＰＶグラフ表示用パネル    Ｙ軸左 表示");
	        for(int i = 0 ; i < 3 ; i++){
	            inc[i]  = (CZPV.getPVGrMax(i) - CZPV.getPVGrMin(i)) / CZPV.PV_Y_SPLIT;
	            leng[i] = CZPV.getPVGrMin(i);
	        }
        } else {
	        for(int i = 5 ; i < 8 ; i++){
	            inc[i-5]  = (CZPV.getPVGrMax(i) - CZPV.getPVGrMin(i)) / CZPV.PV_Y_SPLIT;
	            leng[i-5] = CZPV.getPVGrMin(i);
	        }
        }
// @20131030

        y_pos[0] = -35;
        y_pos[1] = -20;
        y_pos[2] = -5;

        float  h = (float)(rec.height - rec.y) / (float)CZPV.PV_Y_SPLIT;
        for(float y = (float)d.height ; 0 < y ; y-=h){
            g.setColor(CZPV.PV_MEM_COLOR);
            g.drawLine(0,(int)y,d.width,(int)y);

            if (untenflg == true){

//CZSystem.log("CZPVPanelY1View paint","メイン画面ＰＶグラフ表示用パネル    Ｙ軸左 表示２");
	            for(int i = 0 ; i < 3 ; i++){
	                g.setColor(CZPV.PV_COLOR[i]);
	                g.drawString(new String(leng[i] + " " + CZPV.getPVGrUnit(i)),10,(int)y+y_pos[i]); 
	                leng[i]+=inc[i];
	            }
            } else {
	            for(int i = 5 ; i < 8 ; i++){
	                g.setColor(CZPV.PV_COLOR2[i]);
	                g.drawString(new String(leng[i-5] + " " + CZPV.getPVGrUnit(i)),10,(int)y+y_pos[i-5]); 
	                leng[i-5]+=inc[i-5];
	            }
            }
        }
    }
}
