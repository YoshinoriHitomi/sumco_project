package cz;

import javax.swing.JTabbedPane;

/***********************************************************
 *
 *   �W���Ď�����p�p�l��
 * @author  (KPK Co.,Ltd.)
 * @version 1.0 (2003/04/01)
 *
 ***********************************************************/
public class CZCMSControlPanel extends JTabbedPane {

    //�S��&�q�[�^�[
    private CZCMSControlPanel_1 ctlPanel_1  = null;
    //���̑�
    private CZCMSControlPanel_2 ctlPanel_2  = null;

    // ---------- �R���X�g���N�^ ---------------------------
    CZCMSControlPanel(){
        super();

        try{
            setName("CZCMSControlPanel");
            setBackground(java.awt.Color.gray);
            setBounds(20, 65, 800, 250);

            Thread th;

            ctlPanel_1 = new CZCMSControlPanel_1();
            add(ctlPanel_1,"�S��&�q�[�^�[");
            th = new Thread(ctlPanel_1,"CZCMSControlPanel-ctlPanel_1");
            th.start();

            ctlPanel_2 = new CZCMSControlPanel_2();
            add(ctlPanel_2,"���̑�");
            th = new Thread(ctlPanel_2,"CZCMSControlPanel-ctlPanel_2");
            th.start();

//@@            CZSystem.log("CZCMSControlPanel","new");
        }
        catch (Throwable e) {
            CZSystem.handleException(e);
        }
    }
}
