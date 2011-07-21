package cn.ncs.tst;

import org.apache.log4j.Logger;
import org.junit.AfterClass;
import org.junit.BeforeClass;
import org.junit.Test;


public class JreportHelp {

	public static void main(String[] args) {
		int loopCount = 10;
		String loopString = "substr(tt.HOSP_CASE_NO ,1,1) HOSP_CASE_NO_";
		String tmp = "";
		for(int i=1;i<=loopCount;i++){
			tmp = loopString.replaceFirst("\\d", String.valueOf(i));
			tmp = tmp + i+",";
			
			System.out.println(tmp);
		}
	}

	static Logger logger = null;

	@BeforeClass
	public static void init() {
		logger = Logger.getLogger(JreportHelp.class);
	}

	@AfterClass
	public static void over() {
		logger = null;
	}

	@Test
	public void testLogger() throws Exception{
		logger.info("hubin");
	}
	
}
