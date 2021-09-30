package com.dfl.report.ExcelReader;

import java.util.HashMap;

/* *******************************
 * 序列对照表
 * *********************************/

import java.util.Map;

public class SequenceComparisonTable {
	public String getParameterGroup() {
		return ParameterGroup;
	}
	public void setParameterGroup(String parameterGroup) {
		ParameterGroup = parameterGroup;
	}
	public Map<String, String> getValues() {
		return values;
	}
	public void setValues(Map<String, String> fvalues) {
		if(fvalues!=null && fvalues.size()>0) {
			for(Map.Entry<String, String> entry: fvalues.entrySet()) {
				String key = entry.getKey();
				String value = entry.getValue();
				values.put(key, value);
			}
		}
	}
	private String ParameterGroup; //参数组别
    private Map<String,String> values = new HashMap<String,String>(); //B列到P列
}
