package com.sea.quickclick.report.excel.helper;

import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

public class MyDataSet {
	
	private int recordCount;
	private ArrayList<String> lableList = new ArrayList<String>();
	private List<HashMap> dataList = new ArrayList<HashMap>();
	
	public int getDatasetRowsCount(){
		return recordCount;
	}
	
	public HashMap<String, Object> getRecord(int i) {
		return dataList.get(i);
	}
	
	public String getLabel(int i){
		return lableList.get(i);
	}

	public int getRecordCount() {
		return recordCount;
	}

	public void setRecordCount(int recordCount) {
		this.recordCount = recordCount;
	}

	public ArrayList<String> getLableList() {
		return lableList;
	}

	public void setLableList(ArrayList<String> lableList) {
		this.lableList = lableList;
	}

	public List<HashMap> getDataList() {
		return dataList;
	}

	public void setDataList(List<HashMap> dataList) {
		this.dataList = dataList;
	}
	
	
}
