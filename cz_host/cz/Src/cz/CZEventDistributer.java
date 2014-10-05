package cz;


/**
 * Event Distributer
 * @author  (KPK Co.,Ltd.)
 * @version 1.0 (2003/04/01)
 */
public class CZEventDistributer implements Runnable {

    //
    //
    //
    CZEventDistributer(){

    }

    //
    //
    //
    public void run(){

        int i = 0;
        while(true){
//@@            CZSystem.log("CZEventDistributer run","Sleep");
            CZSystem.sleep(10000);
        }
    }
} 
