package cz;

import java.awt.Dimension;
import java.awt.Graphics;
import java.awt.Rectangle;

import javax.swing.JPanel;

/**
 *   メイン画面ＰＶグラフ表示用パネル
 * @author  (KPK Co.,Ltd.)
 * @version 1.0 (2003/04/01)
 */
public class CZPVPanelXView extends JPanel {
    private Rectangle rec = null;
    
    CZPVPanelXView(){
        super();

        try{
            setName("CZPVPanelXView");
            setLayout(null);
            setBackground(CZPV.PV_BACK_COLOR);

//@@            CZSystem.log("CZPVPanelXView","new");
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

        int inc;
        String unit = null;

        if(CZPV.getPVGrTimeFlag()){
            inc = CZPV.TIME_SCALE_ST * CZPV.getPVGrTimeScale();
            unit = new String("min");
        }
        else{
            inc = CZPV.LENGTH_SCALE_ST * CZPV.getPVGrLengthScale();
            unit = new String("mm");
        }


        g.setColor(CZPV.PV_MEM_COLOR);

        float w = (float)(rec.width - rec.x) / (float)CZPV.PV_X_SPLIT;

        int x_scale = 0;
        for(float x = 0 ; x < d.width ; x+=w){
            g.drawLine((int)x,0,(int)x,rec.height / 4);

            g.drawString(new String(x_scale + unit),(int)x+5,rec.height / 2);
            x_scale += inc;
        }
    }
}
