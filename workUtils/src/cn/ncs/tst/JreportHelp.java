package cn.ncs.tst;

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

}
