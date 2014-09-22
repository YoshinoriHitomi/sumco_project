package cz;

import java.awt.Dimension;
import java.awt.Graphics;
import java.awt.Rectangle;

import javax.swing.JPanel;

/*
 *   メイン画面ＰＶグラフ表示用パネル
 * @author  (KPK Co.,Ltd.)
 * @version 1.0 (2003/04/01)
 * 2008.10.22 H.Nagamine 初期画面表示状態変更
 * Update 2013/10/30 表示切り替え機能 (@20131030)
 */
public class CZPVPanelMainView extends JPanel {

    private Rectangle rec = null;

    private int pvx_db[];
    private int pvy_db[][];

    private int pvx_now[];
    private int pvy_now[][];

    private boolean untenflg = true;	// @20131030
    
    CZPVPanelMainView(){
        super();

        try{
            setName("CZPVPanelMainView");
            setLayout(null);
            setBackground(CZPV.PV_BACK_COLOR);

//@@            CZSystem.log("CZPVPanelMainView","new");
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
    // データベースから読み込んだPVデータの座標変換
    //
// chg start 2008.10.22
//     public void setDBData(){
    public int setDBData(){
// chg end 2008.10.22
        int count = CZPV.getPVDataDBCount();

//@@        CZSystem.log("CZPVPanelMainView setDBData","Count[" + count + "]");

        if(2 > count) return -1; 
        pvx_db = new int[count];
        pvy_db = new int[10][count];	// @20131030

        Dimension d = getSize(null);

        float y1,y2;

// @20131030
        untenflg = CZSystem.untenView();
        if (untenflg == true){
	        for(int i = 0 ; i < 5 ; i++){
	            y1 = (float)(rec.height - rec.y) / (CZPV.getPVGrMax(i) - CZPV.getPVGrMin(i)); 
	            for(int j = 0 ; j < count ; j++){
	                y2 = y1 * (CZPV.getPVDataDB(j,i+2) - CZPV.getPVGrMin(i));
	                pvy_db[i][j] = (int)(d.height - y2);
	            }
	        }
        } else {
	        for(int i = 5 ; i < 10 ; i++){
	            y1 = (float)(rec.height - rec.y) / (CZPV.getPVGrMax(i) - CZPV.getPVGrMin(i)); 
	            for(int j = 0 ; j < count ; j++){
	                y2 = y1 * (CZPV.getPVDataDB(j,i+2) - CZPV.getPVGrMin(i));
	                pvy_db[i][j] = (int)(d.height - y2);
	            }
	        }
        }
// @20131030

        float x1,x2;

        if(CZPV.getPVGrTimeFlag()){
            x1 = (float)(rec.width - rec.x) / 
                    (float)(CZPV.PV_X_SPLIT* CZPV.TIME_SCALE_ST * CZPV.getPVGrTimeScale() * (float)60.0);
            for(int j = 0 ; j < count ; j++){
                pvx_db[j] = (int)(x1 * CZPV.getPVDataDB(j,0));
            }
        }
        else {
            x1 = (float)(rec.width - rec.x) / 
                (float)(CZPV.PV_X_SPLIT * CZPV.LENGTH_SCALE_ST * CZPV.getPVGrLengthScale());
            for(int j = 0 ; j < count ; j++){
                pvx_db[j] = (int)(x1 * CZPV.getPVDataDB(j,1));
            }
        }
// add start 2008.10.22
        return count;
// add end 2008.10.22
    }

    //
    // 現在のＰＶデータの座標変換
    //
// chg start 2008.10.22
//    public void setNOWData(){
    public int setNOWData(){
// chg end 2008.10.22
        int count = CZPV.getPVDataUseCount();

//@@        CZSystem.log("CZPVPanelMainView setNOWData","Count 1 [" + count + "]");

        if(2 > count) return -1; 
        pvx_now = new int[count];
        pvy_now = new int[10][count];	// @20131030

        Dimension d = getSize(null);

        float y1,y2;

// @20131030
        untenflg = CZSystem.untenView();
        if (untenflg == true){
	        for(int i = 0 ; i < 5 ; i++){
	            y1 = (float)(rec.height - rec.y) / (CZPV.getPVGrMax(i) - CZPV.getPVGrMin(i)); 
	            for(int j = 0 ; j < count ; j++){
	                y2 = y1 * (CZPV.getPVDataUse(j,i+2) - CZPV.getPVGrMin(i));
	                pvy_now[i][j] = (int)(d.height - y2);
	            }
	        }
        } else {
	        for(int i = 5 ; i < 10 ; i++){
	            y1 = (float)(rec.height - rec.y) / (CZPV.getPVGrMax(i) - CZPV.getPVGrMin(i)); 
	            for(int j = 0 ; j < count ; j++){
	                y2 = y1 * (CZPV.getPVDataUse(j,i+2) - CZPV.getPVGrMin(i));
	                pvy_now[i][j] = (int)(d.height - y2);
	            }
	        }
        }
// @20131030

        float x1,x2;
            
        if(CZPV.getPVGrTimeFlag()){
            x1 = (float)(rec.width - rec.x) / 
                (float)(CZPV.PV_X_SPLIT* CZPV.TIME_SCALE_ST * CZPV.getPVGrTimeScale() * (float)60.0);
            for(int j = 0 ; j < count ; j++){
                pvx_now[j] = (int)(x1 * CZPV.getPVDataUse(j,0));
            }
        }
        else {
            x1 = (float)(rec.width - rec.x) / 
                (float)(CZPV.PV_X_SPLIT * CZPV.LENGTH_SCALE_ST * CZPV.getPVGrLengthScale());
            for(int j = 0 ; j < count ; j++){
                pvx_now[j] = (int)(x1 * CZPV.getPVDataUse(j,1));
            }
        }
// add start 2008.10.22
        return count;
// add end 2008.10.22
    }


    //
    // 
    //
    public void paint(Graphics g){

        Dimension d = getSize(null);
        g.setColor(CZPV.PV_BACK_COLOR);
        g.fillRect(0,0,d.width,d.height);

        // グラフ目盛の描画
        g.setColor(CZPV.PV_MEM_COLOR);

        float w = (float)(rec.width - rec.x) / (float)CZPV.PV_X_SPLIT;
        float sp1 = w / (float)2 ;
        float sp2 = w / (float)4 ;
        for(float x = 0 ; x < d.width ; x+=w){
            g.setColor(CZPV.PV_MEM_SP2_COLOR);
            g.drawLine((int)(x+sp2),0,(int)(x+sp2),d.height);
            g.drawLine((int)(x-sp2),0,(int)(x-sp2),d.height);
            g.setColor(CZPV.PV_MEM_SP1_COLOR);
            g.drawLine((int)(x+sp1),0,(int)(x+sp1),d.height);
            g.setColor(CZPV.PV_MEM_COLOR);
            g.drawLine((int)x,0,(int)x,d.height);
        }

        float h  = (float)(rec.height - rec.y) / (float)CZPV.PV_Y_SPLIT;
        sp1 = h / (float)2 ;
        sp2 = h / (float)4 ;
        for(float y = (float)d.height ; 0 < y ; y-=h){
            g.setColor(CZPV.PV_MEM_SP2_COLOR);
            g.drawLine(0,(int)(y-sp2),d.width,(int)(y-sp2));
            g.drawLine(0,(int)(y+sp2),d.width,(int)(y+sp2));

            g.setColor(CZPV.PV_MEM_SP1_COLOR);
            g.drawLine(0,(int)(y-sp1),d.width,(int)(y-sp1));

            g.setColor(CZPV.PV_MEM_COLOR);
            g.drawLine(0,(int)y,d.width,(int)y);
        }


        // データベースより読み込んだPVの描画
        if(2 < CZPV.getPVDataDBCount()){
            try{
                setDBData();
// @20131030
		        untenflg = CZSystem.untenView();
		        if (untenflg == true){

	                for(int i = 0 ; i < 5 ; i++){
	                    g.setColor(CZPV.PV_COLOR[i]);
	                    g.drawPolyline(pvx_db,pvy_db[i],pvx_db.length);
	                }
                } else {
	                for(int i = 5 ; i < 10 ; i++){
	                    g.setColor(CZPV.PV_COLOR2[i]);
	                    g.drawPolyline(pvx_db,pvy_db[i],pvx_db.length);
	                }
                }
// @20131030

            }
            catch (Throwable e) {
                CZSystem.log("CZPVPanelMainView paint","Data Error 1" + e);
//@@                System.out.println(e);
            }
        }


        // 現在ＰＶの描画
        if(2 < CZPV.getPVDataUseCount()){
            try{
                setNOWData();

// @20131030
//		        untenflg = CZSystem.untenView();
		        if (untenflg == true){
	                for(int i = 0 ; i < 5 ; i++){
	                    g.setColor(CZPV.PV_COLOR[i]);
	                    g.drawPolyline(pvx_now,pvy_now[i],pvx_now.length);
	                }
                } else {
	                for(int i = 5 ; i < 10 ; i++){
	                    g.setColor(CZPV.PV_COLOR2[i]);
	                    g.drawPolyline(pvx_now,pvy_now[i],pvx_now.length);
	                }
                }
// @20131030

            }
            catch (Throwable e) {
                CZSystem.log("CZPVPanelMainView paint","Data Error 2" + e);
//@@                System.out.println(e);
            }
        }
    }
}
