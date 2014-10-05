package cz;

import java.io.FileInputStream;
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.sql.Statement;
import java.util.Properties;

/**
 * データベースよりPVをReadする。
 * @author  (KPK Co.,Ltd.)
 * @version 1.0 (2003/04/01)
 * Update 2013/10/30 表示切り替え機能 (@20131030)
 */
public class CZPVDBReader implements Runnable {

    private String dbName_      = null;
    private String viewName_    = null;
    private int procunq_        = -1;
    private int pv1_            = -1;
    private int pv2_            = -1;
    private int pv3_            = -1;
    private int pv4_            = -1;
    private int pv5_            = -1;
    private int pv6_            = -1;	// @20131030
    private int pv7_            = -1;	// @20131030
    private int pv8_            = -1;	// @20131030
    private int pv9_            = -1;	// @20131030
    private int pv10_            = -1;	// @20131030

    //
    //
    //
    CZPVDBReader(String db,String v,int p,int v1,int v2,int v3,int v4,int v5,int v6,int v7,int v8,int v9,int v10){	// @20131030
        dbName_     = db;
        viewName_   = v;
        procunq_    = p;

        pv1_    = v1;
        pv2_    = v2;
        pv3_    = v3;
        pv4_    = v4;
        pv5_    = v5;
        pv6_    = v6;	// @20131030
        pv7_    = v7;	// @20131030
        pv8_    = v8;	// @20131030
        pv9_    = v9;	// @20131030
        pv10_    = v10;	// @20131030
        
//@@        System.out.println("CZPVDBReader---------->pv:"+pv1_+":"+pv2_+":"+pv3_+":"+pv4_+":"+pv5_);
    }

    //
    //
    //
    public void run(){

//@@        CZSystem.log("CZPVDBReader","Start !!");

        if(null == dbName_){
            CZSystem.exit(-1,new String("CZPVDBReader run() Bad DB NAME[" + dbName_ + "]"));
        };
//@@        CZSystem.log("CZPVDBReader","dbName OK !!" + dbName_);

        if(null == viewName_){
            CZSystem.exit(-1,new String("CZPVDBReader run() Bad VIEW NAME[" + viewName_ + "]"));
        };
//@@        CZSystem.log("CZPVDBReader","viewName OK!!" + viewName_);

        if(0 > procunq_){
            CZSystem.exit(-1,new String("CZPVDBReader run() Bad UNIQUE PROCESS [" + procunq_ + "]"));
        };
//@@        CZSystem.log("CZPVDBReader","procunq OK!!" + procunq_);

        if((1 > pv1_) || (CZPV.PV_MAX_LENGTH < pv1_)){
            CZSystem.exit(-1,new String("CZPVDBReader run() Bad PV1 NAME [" + pv1_ + "]"));
        };
//@@        CZSystem.log("CZPVDBReader","pv1 OK!!" + pv1_);

        if((1 > pv2_) || (CZPV.PV_MAX_LENGTH < pv2_)){
            CZSystem.exit(-1,new String("CZPVDBReader run() Bad PV2 NAME [" + pv2_ + "]"));
        };
//@@        CZSystem.log("CZPVDBReader","pv2 OK!!" + pv2_);

        if((1 > pv3_) || (CZPV.PV_MAX_LENGTH < pv3_)){
            CZSystem.exit(-1,new String("CZPVDBReader run() Bad PV3 NAME [" + pv3_ + "]"));
        };
//@@        CZSystem.log("CZPVDBReader","pv3 OK!!" + pv3_);

        if((1 > pv4_) || (CZPV.PV_MAX_LENGTH < pv4_)){
            CZSystem.exit(-1,new String("CZPVDBReader run() Bad PV4 NAME [" + pv4_ + "]"));
        };
//@@        CZSystem.log("CZPVDBReader","pv4 OK!!" + pv4_);

        if((1 > pv5_) || (CZPV.PV_MAX_LENGTH < pv5_)){
            CZSystem.exit(-1,new String("CZPVDBReader run() Bad PV5 NAME [" + pv5_ + "]"));
        };
//@@        CZSystem.log("CZPVDBReader","pv5 OK!!" + pv5_);

// @20131030
        if((1 > pv6_) || (CZPV.PV_MAX_LENGTH < pv6_)){
            CZSystem.exit(-1,new String("CZPVDBReader run() Bad PV6 NAME [" + pv6_ + "]"));
        };
//@@        CZSystem.log("CZPVDBReader","pv6 OK!!" + pv6_);

        if((1 > pv7_) || (CZPV.PV_MAX_LENGTH < pv7_)){
            CZSystem.exit(-1,new String("CZPVDBReader run() Bad PV7 NAME [" + pv7_ + "]"));
        };
//@@        CZSystem.log("CZPVDBReader","pv7 OK!!" + pv7_);

        if((1 > pv8_) || (CZPV.PV_MAX_LENGTH < pv8_)){
            CZSystem.exit(-1,new String("CZPVDBReader run() Bad PV8 NAME [" + pv8_ + "]"));
        };
//@@        CZSystem.log("CZPVDBReader","pv8 OK!!" + pv8_);

        if((1 > pv9_) || (CZPV.PV_MAX_LENGTH < pv9_)){
            CZSystem.exit(-1,new String("CZPVDBReader run() Bad PV9 NAME [" + pv9_ + "]"));
        };
//@@        CZSystem.log("CZPVDBReader","pv9 OK!!" + pv9_);

        if((1 > pv10_) || (CZPV.PV_MAX_LENGTH < pv10_)){
            CZSystem.exit(-1,new String("CZPVDBReader run() Bad PV5 NAME [" + pv10_ + "]"));
        };
//@@        CZSystem.log("CZPVDBReader","pv10 OK!!" + pv10_);
// @20131030

//@@        CZSystem.log("CZPVDBReader","dbRead Call !!");
        int ret = dbRead();
        if(0 > ret){
            CZSystem.exit(-1,"CZPVDBReader run() DATABASE ERROR !!");
        }

//@@        CZSystem.log("CZPVDBReader","CZEventSender Call !!");
        CZEventSender.sendData(this,CZEventCL.RO_CHANGE);
//@@        CZSystem.log("CZPVDBReader","END !!");
    }


    //
    //
    //
    private int dbRead(){

        String dbUrl        = null;     //@@
        String host         = null;
        String user         = null;
        String passwd       = null;
        String port         = null;

        Properties pr       = new Properties();
        Connection conn     = null;
        Statement  sqlstmt  = null;
        ResultSet  rs       = null;
        String     sql      = null;

        try{
            Properties prop =  new Properties();
            FileInputStream pros = new FileInputStream(CZSystemDefine.PROPERTY_FILE);
            prop.load(pros);

            prop.list(System.out);

            host    = prop.getProperty("HOST");
            user    = prop.getProperty("USER");
            passwd  = prop.getProperty("PASSWD");
//@@            port    = prop.getProperty("PORT");
            dbUrl   = prop.getProperty("DB_URL");               //@@

            pr.put("USER",user);
            pr.put("PASSWORD",passwd);
        }
        catch( Exception e){
            CZSystem.exit(-1,"dbRead() NO Propertie File");
        }


        int i = 0;
        sql = new String("SELECT p_time,data5,data" + 		// @20131030
                        pv1_ + ",data" + pv2_ + ",data" + 	// @20131030
                        pv3_ + ",data" + pv4_ + ",data" + 	// @20131030
                        pv5_ + ",data" + pv6_ + ",data" + 	// @20131030
                        pv7_ + ",data" + pv8_ + ",data" + 	// @20131030
                        pv9_ + ",data" + pv10_ + 			// @20131030
                        " FROM " + dbName_ + "." + viewName_ + " WHERE p_renban = " + procunq_ +
                        " ORDER BY p_time");
        CZSystem.log("CZPVDBReader","dbRead() SQL[" + sql + "]");
        // ドライバをインスタンス化します。
        try{
            DriverManager.registerDriver(new oracle.jdbc.driver.OracleDriver());
        }
        catch (Throwable e) {
            CZSystem.log("CZPVDBReader","ERROR: failed to load JDBC driver.");
            CZSystem.handleException(e);
            return -1;
        }

        // ドライバ接続
        try{
            conn = DriverManager.getConnection(dbUrl,user,passwd);               //@@

        }
        catch (SQLException e) {
            CZSystem.log("CZPVDBReader","ERROR: failed to connect!");
            CZSystem.handleException(e);
            return -1;
        }

        try{
            sqlstmt = conn.createStatement() ;
        }
        catch(SQLException e){
            closeConnect(conn);
            CZSystem.log("CZPVDBReader","ERROR: createStatement or database");
            return -1;
        }

        try{
            rs = sqlstmt.executeQuery(sql) ;
            for(i = 0 ; rs.next() ; i++){
                CZPV.addPVDataDB(
                    rs.getFloat(1),rs.getFloat(2),
                    rs.getFloat(3),rs.getFloat(4),
                    rs.getFloat(5),rs.getFloat(6),rs.getFloat(7),rs.getFloat(8),rs.getFloat(9),rs.getFloat(10),rs.getFloat(11),rs.getFloat(12));	// @20131030
            } // for end
            CZSystem.log("CZPVDBReader","SELECT Count:" + i);
        }
        catch( SQLException e ){
            CZSystem.log("CZPVDBReader","ERROR: Select failed.");
        }

        try{
            if(null != rs) rs.close();             //@@
            sqlstmt.close();        //@@
        }
        catch (SQLException e){
            CZSystem.log("CZPVDBReader","dbRead () ERROR: Close ResultSet or Statement");
        }
        closeConnect(conn);
        return i;
    }   


    //
    //
    //
    private boolean closeConnect(Connection c){

        try{
            c.close();
        }
        catch (SQLException e){
            CZSystem.log("CZPVDBReader","ERROR: Close Connection");
            return false;
        }
        return true;
    }
}
