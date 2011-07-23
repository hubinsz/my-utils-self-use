package com.sea.quickclick.report.excel.helper;

import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

public class MyDataSet {
	
	private int recordCount;
	private Map<String, Object> lableMap = new HashMap<String, Object>();
	private List<HashMap> dataList = new ArrayList<HashMap>();
	
	public int getDatasetRowsCount(){
		return recordCount;
	}
	
	public HashMap<String, Object> getRecord(int i) {
		return dataList.get(i);
	}
}
