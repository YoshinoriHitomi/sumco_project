package cz;

import javax.swing.JPanel;

/***********************************************************
 *
 *   ���C����ʗp�p�l��
 * @author  (KPK Co.,Ltd.)
 * @version 1.0 (2003/04/01)
 *
 ***********************************************************/
public class CZMainPanel extends JPanel {

    CZMainPanel(){
        super();
//@@        CZSystem.log("CZMainPanel","new");

        try{
            setName("JAppletContentPane");
            setLayout(null);

        }
        catch (Throwable e) {
          CZSystem.handleException(e);
        }
    }
}
