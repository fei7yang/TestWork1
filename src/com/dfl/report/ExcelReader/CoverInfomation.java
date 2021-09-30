package com.dfl.report.ExcelReader;

public class CoverInfomation {

	private String linebody; // 线体
	private String Edition;// 版次
	private String Factory; // 工厂
	private String filecode; // 工厂

	public String getFilecode() {
		return filecode;
	}

	public void setFilecode(String filecode) {
		this.filecode = filecode;
	}

	public String getFactory() {
		return Factory;
	}

	public void setFactory(String factory) {
		Factory = factory;
	}

	public String getLinebody() {
		return linebody;
	}

	public void setLinebody(String linebody) {
		this.linebody = linebody;
	}

	public String getEdition() {
		return Edition;
	}

	public void setEdition(String edition) {
		Edition = edition;
	}
}
