package com.ncs.jreport.helper;

import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.sql.Statement;

import org.apache.log4j.Logger;
import org.junit.AfterClass;
import org.junit.BeforeClass;
import org.junit.Test;

public class JreportHelper {

	static String dbUrl = "jdbc:oracle:thin:@192.168.82.43:1521:ndrs";
	static String userName = "ndradmin";
	static String password = "ndradmin";
	static Connection connection = null;
	static Statement statement = null;
	static PreparedStatement preparedStatement = null;
	static ResultSet resultSet = null;

	static Logger logger = null;

	// configuration
	static String code_id = "TRA_YES_NO_ND";
	static String code_desc_type = "CODE_DESC";
	//static String code_desc_type = "SHORT_DESC";
	static String reg_field = "REHAB_FIM_SCORE_ST";
	static String sectionId = "630";
	
	//end configuration

	@BeforeClass
	public static void init() throws Exception {
		Class.forName("oracle.jdbc.driver.OracleDriver");
		logger = Logger.getLogger(JreportHelper.class);
	}

	@Test
	public void run001getCode() throws Exception {
		int recordNum = 0;//test hubin 2011-07-21
		connection = DriverManager.getConnection(dbUrl, userName, password);
		statement = connection.createStatement();
		String sql = "select count(1) from tbl_code_int_grp t where t.CODETYPE_ID = '"
				+ code_id + "'";
		resultSet = statement.executeQuery(sql);
		while (resultSet.next()) {
			recordNum = resultSet.getInt(1);
		}
		//logger.info(recordNum);
		StringBuffer stringBuffer = new StringBuffer();
		stringBuffer.append("\r\nWITH TMP_").append(sectionId).append("_");
		stringBuffer.append(code_id).append(" AS ( ").append("\r\n");
		stringBuffer.append("SELECT \r\nt1.codetype_id, \r\n");
		String loopNumTmp = "";
		//before from
		for(int i=1;i<=recordNum;i++){
			loopNumTmp = String.valueOf(i);
			stringBuffer.append("t")
			.append(loopNumTmp)
			.append(".code_id ")
			.append(code_id).append("_").append(sectionId)
			.append("_")
			.append(loopNumTmp)
			.append("_ID")
			.append(",  t").append(loopNumTmp).append(".").append(code_desc_type)
			.append("  ").append(code_id).append("_").append(sectionId).append("_").append(loopNumTmp);
			if(i==recordNum){
				stringBuffer.append("_TT \r\n");
			}else{
				stringBuffer.append("_TT, \r\n");
			}
		}
		stringBuffer.append("FROM \r\n");
		//after from
		for(int i=1;i<=recordNum;i++){
			loopNumTmp = String.valueOf(i);
			stringBuffer.append("tbl_code_int_grp t")
			.append(loopNumTmp);
			if(i==recordNum){
				stringBuffer.append(" \r\n");
			}else{
				stringBuffer.append(", \r\n");
			}
		}
		
		stringBuffer.append("WHERE 1=1 \r\n");
		//after where
		//t1.codetype_id='D_RESI' and t1.code_seq=1 
		for(int i=1;i<=recordNum;i++){
			loopNumTmp = String.valueOf(i);
			stringBuffer.append("and t")
			.append(loopNumTmp).append(".codetype_id = '"+code_id+"'").append(" and ").append("t").append(loopNumTmp)
			.append(".code_seq=").append(loopNumTmp).append(" \r\n");
		}
		
		stringBuffer.append(")");
		logger.info(stringBuffer.toString());
		
		//case when
		StringBuffer stringBufferCaseWhen = new StringBuffer("\r\n");
		for(int i=1;i<=recordNum;i++){
			loopNumTmp = String.valueOf(i);
			stringBufferCaseWhen.append("CASE WHEN tblreg.")
			.append(reg_field)
			.append(" = ")
			.append("TMP_").append(sectionId).append("_").append(code_id)
			.append(".").append(code_id).append("_").append(sectionId).append("_").append(loopNumTmp).append("_ID")
			.append(" THEN 'X' ELSE '' END ")
			.append(code_id).append("_").append(sectionId).append("_").append(loopNumTmp).append("_BX, ")
			.append("TMP_").append(sectionId).append("_").append(code_id).append(".")
			.append(code_id).append("_").append(sectionId).append("_").append(loopNumTmp).append("_TT ");
			if(i==recordNum){
				stringBufferCaseWhen.append("\r\n");
			}else{
				stringBufferCaseWhen.append(", \r\n");
			}
		}
		
		logger.info(stringBufferCaseWhen.toString());
		
		//join
		StringBuffer stringBufferJoin = new StringBuffer("\r\n");
		stringBufferJoin.append("LEFT JOIN ")
		.append("TMP_").append(sectionId).append("_").append(code_id)
		.append(" ON ")
		.append("TMP_").append(sectionId).append("_").append(code_id).append(".CODETYPE_ID='"+code_id+"'");
		logger.info(stringBufferJoin.toString());
		
		free(resultSet, preparedStatement, connection);
	}

	@Test
	public void run002SplitField(){
		int splitTimes = 13;
		String db_field = "SCDF_INCIDENT_NO";
		StringBuffer stringBuffer = new StringBuffer("\r\n");
		stringBuffer.append("tblreg.").append(db_field).append(",\r\n");
		String ii= "";
		for(int i=1;i<=splitTimes;i++){
			ii=String.valueOf(i);
			stringBuffer.append("SUBSTR(tblreg.")
			.append(db_field).append(",").append(ii).append(",1) ")
			.append(db_field).append("_").append(ii).append(",\r\n");
		}
		logger.info(stringBuffer.toString());
	}
	@Test
	public void run003SplitFieldOther(){
		int splitTimes = 13;
		String db_field = "SCDF_INCIDENT_NO";
		StringBuffer stringBuffer = new StringBuffer("\r\n");
		stringBuffer.append("tblreg.").append(db_field).append(",\r\n");
		String ii= "";
		for(int i=1;i<=splitTimes;i++){
			ii=String.valueOf(i);
			stringBuffer.append("SUBSTR(tblreg.")
			.append(db_field).append(",").append(ii).append(",1) ")
			.append(db_field).append("_").append(ii).append(",\r\n");
		}
		logger.info(stringBuffer.toString());
	}
	
	@AfterClass
	public static void allover() {
		logger = null;
	}

	public void free(ResultSet resultSet, Statement statement,
			Connection connection) {
		try {
			if (resultSet != null) {
				resultSet.close();
			}
		} catch (SQLException e) {
			e.printStackTrace();
		} finally {
			try {
				if (statement != null) {
					statement.close();
				}
			} catch (SQLException e) {
				e.printStackTrace();
			} finally {
				if (connection != null) {
					try {
						connection.close();
					} catch (SQLException e) {
						e.printStackTrace();
					}
				}
			}
		}
	}

}
