package cz;

import java.util.Vector;

/**
 *  Queue
 * @author  (KPK Co.,Ltd.)
 * @version 1.0 (2003/04/01)
 */
public class CZSystemQueue {

    private Vector theQueue;
    private int    maxQue;

    //
    //  コンストラクタ
    //
    public CZSystemQueue() {
        theQueue = new Vector();
        maxQue   = 0;
    }   

    public CZSystemQueue(int _maxQue) { 
        theQueue = new Vector();    
        maxQue   = _maxQue;
    }   

    //
    //  Queueの大きさを設定する。
    //
    public synchronized void setMaxQue(int _maxQue) {   
        maxQue = _maxQue;
    }   

    //
    // データをPutする。
    //
	@SuppressWarnings("unchecked")
    public synchronized void put(Object toPut) {    
        if (maxQue > 0) {
            while (size() >= maxQue) {
                Object oldObject = get();
            }
        }

        theQueue.addElement(toPut); 
        notify();
    }

    //
    //  データを取得する。
    //
    public synchronized Object get() {  
        Object  found = peekAtHead();   
        if (found != null) {    
            theQueue.removeElementAt(0);    
        }   
        return found;
    }   

    public synchronized Object waitObject() throws InterruptedException {   
        while (isEmpty()) { 
            wait(); 
        }   
        return get();   
    }   

    public synchronized Object peekAtHead() {   
        if (theQueue.isEmpty()) {
            return null;
        }
        return theQueue.elementAt(0);
    }

    public boolean isEmpty() {
        return theQueue.isEmpty();
   }
   
   public int size() {
       return theQueue.size();
   }
}
