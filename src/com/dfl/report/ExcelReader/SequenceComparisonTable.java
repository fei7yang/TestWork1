package com.dfl.report.ExcelReader;

import java.util.HashMap;

/* *******************************
 * ���ж��ձ�
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
	private String ParameterGroup; //�������
    private Map<String,String> values = new HashMap<String,String>(); //B�е�P��
}
