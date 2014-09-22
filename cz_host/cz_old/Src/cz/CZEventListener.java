package cz;

import java.util.EventListener;

/**
 * Event Listener
 * @author  (KPK Co.,Ltd.)
 * @version 1.0 (2003/04/01)
 */
public interface CZEventListener extends EventListener {
    public void arrival(CZEventCL e);
}
