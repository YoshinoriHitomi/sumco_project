package cz;

import java.util.Vector;

/**
 * Event Sender
 * @author  (KPK Co.,Ltd.)
 * @version 1.0 (2003/04/01)
 */
public class CZEventSender {

    private static Vector listeners = new Vector();
    
    //
    //
    //
	@SuppressWarnings("unchecked")
    public static boolean addCZEventListener(CZEventListener l){
        try{
            listeners.addElement(l);
            return true;
        }
        catch (Throwable e) {
            CZSystem.handleException(e);
            return false;
        }
    }

    //
    //
    //
    public static synchronized boolean  removeCZEventListener(CZEventListener l){
        try{
            listeners.removeElement(l);
            return true;
        }
        catch (Throwable e) {
            CZSystem.handleException(e);
            return false;
        }
    }


    //
    //
    //
    public static synchronized boolean  sendData(Object obj,int event){
//@@        CZSystem.log("CZEventSender sendData","targets[" + listeners.size() + "]");

        int i;

        try{
            CZEventCL ev = new CZEventCL(obj,event);
            Vector targets = (java.util.Vector)listeners.clone();
            for(i = 0 ; i < targets.size() ; i++) {
                CZEventListener listen = (CZEventListener)targets.elementAt(i);
                listen.arrival(ev);
            }
            return true;
        }
        catch (Throwable e) {
            CZSystem.handleException(e);
            return false;
        }
    }
}
