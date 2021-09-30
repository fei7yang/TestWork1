package com.dfl.report.util;

import com.teamcenter.rac.kernel.TCComponentItemRevision;

public class GenerateReportInfo {
	private boolean isExist;
	private boolean isgoon;
	private String action;
	private TCComponentItemRevision meDocument;
	private String DFL9_process_type;//文档类型
	private String DFL9_process_file_type;//文件类型
	private String meDocumentName;//文件名称
	private boolean flag = false;//用于区分是否需要通过名称校验
	private TCComponentItemRevision project_ids;


	
	public GenerateReportInfo(boolean isExist,boolean isgoon,
			String action,TCComponentItemRevision meDocument,
			String DFL9_process_type,String DFL9_process_file_type){
		this.isExist = isExist;
		this.isgoon = isgoon;
		this.action = action;
		this.meDocument = meDocument;
		this.DFL9_process_file_type = DFL9_process_file_type;
		this.DFL9_process_type = DFL9_process_type;
	}
	public GenerateReportInfo(){
		
	}
	
	public boolean isExist() {
		return isExist;
	}

	public void setExist(boolean isExist) {
		this.isExist = isExist;
	}

	public boolean isIsgoon() {
		return isgoon;
	}

	public void setIsgoon(boolean isgoon) {
		this.isgoon = isgoon;
	}

	public String getAction() {
		return action;
	}

	public void setAction(String action) {
		this.action = action;
	}

	public TCComponentItemRevision getMeDocument() {
		return meDocument;
	}

	public void setMeDocument(TCComponentItemRevision meDocument) {
		this.meDocument = meDocument;
	}

	public String getDFL9_process_type() {
		return DFL9_process_type;
	}

	public void setDFL9_process_type(String DFL9_process_type) {
		this.DFL9_process_type = DFL9_process_type;
	}

	public String getDFL9_process_file_type() {
		return DFL9_process_file_type;
	}

	public void setDFL9_process_file_type(String DFL9_process_file_type) {
		this.DFL9_process_file_type = DFL9_process_file_type;
	}
	public String getmeDocumentName() {
		return meDocumentName;
	}

	public void setmeDocumentName(String meDocumentName) {
		this.meDocumentName = meDocumentName;
	}
	public boolean isFlag() {
		return flag;
	}
	public void setFlag(boolean flag) {
		this.flag = flag;
	}
	public TCComponentItemRevision getProject_ids() {
		return project_ids;
	}
	public void setProject_ids(TCComponentItemRevision project_ids) {
		this.project_ids = project_ids;
	}
}
