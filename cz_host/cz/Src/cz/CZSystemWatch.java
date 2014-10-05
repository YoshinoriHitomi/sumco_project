package cz;

import czclass.CZRaidStatus;

/*******************************************************************************
 *
 *   システム状態を監視する
 *       定周期ですが処理時間の分、遅れが発生します
 * @author  (KPK Co.,Ltd.)
 * @version 1.0 (2003/04/01)
 *
 *******************************************************************************/
public class CZSystemWatch implements Runnable {
    private Runtime rt = null;
    
    private final int RAID_TIME =  3600 * 2;
    private int   raid_count    = 0;
    CZRaidStatus  raid1         = null;
    CZRaidStatus  raid5         = null;

    private final int GC_TIME   = 10;
    private int   gc_count      = 0;

    //
    //
    CZSystemWatch(){
        rt = Runtime.getRuntime();
    }

    //
    //
    public void run(){

//@@        CZSystem.log("CZSystemWatch","----- START !! -----");

        while(true){
//            CZSystem.log("CZSystemWatch","----- Active !! -----");

            // RAID Check
/*****************************
            raidWatch();
******************************/

            // GC START
            gcWatch();

            CZSystem.sleep(1000);
        }
    }

    //
    //
    private boolean raidWatch(){

        if(0 >= raid_count){
            raid_count = RAID_TIME;

            // RAID Check
            raid1 = CZSystem.CZRaidGetStatus(0,0);
            if(null != raid1)
            CZSystem.log("CZSystemWatch RAID"," RAID1[ " + raid1.getStatus() + "] Log[" + raid1.getLog() + "]");

            raid5 = CZSystem.CZRaidGetStatus(0,1);
            if(null != raid5)
            CZSystem.log("CZSystemWatch RAID"," RAID5[ " + raid5.getStatus() + "] Log[" + raid5.getLog() + "]");

            else raid_count = 0;

            return true;
        }
        raid_count--;
        return false;
    }

    //
    //
    private boolean gcWatch(){

        if(0 >= gc_count){
            gc_count = GC_TIME;
/*
            CZSystem.log("CZSystemWatch GC",
            "FreeMemory [" + rt.freeMemory() + "]  TotalMemory [" + rt.totalMemory() + "]");
*/
//            System.out.println(Runtime.getRuntime().freeMemory());
            System.gc();
//            System.out.println(Runtime.getRuntime().freeMemory() + "  GC FREE!!");
/*
            CZSystem.log("CZSystemWatch   ",
            "FreeMemory [" + rt.freeMemory() + "]  TotalMemory [" + rt.totalMemory() + "]");
*/
            return true;
        }

        gc_count--;
        return false;
    }
}
