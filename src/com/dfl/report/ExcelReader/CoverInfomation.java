package com.dfl.report.ExcelReader;

public class CoverInfomation {

	private String linebody; // ����
	private String Edition;// ���
	private String Factory; // ����
	private String filecode; // ����

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
