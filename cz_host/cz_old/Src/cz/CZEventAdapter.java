package cz;


/**
 *  Event Adapter
 * @author  (KPK Co.,Ltd.)
 * @version 1.0 (2003/04/01)
 */
public class CZEventAdapter implements CZEventListener {
    private CZEventCL event     = null;
    private CZSystemQueue que   = null;

    //
    //
    //
    CZEventAdapter(CZSystemQueue q){
        que = q;
    }

    //
    //
    //
    public void arrival(CZEventCL e){
        event = e;
        que.put(event);
    }
}
