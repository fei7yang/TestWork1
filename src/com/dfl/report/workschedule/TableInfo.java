package com.dfl.report.workschedule;

public class TableInfo {

	
	private String name;
	private String page;
	private boolean canEdit;
	
	public TableInfo(String name,String page,boolean canEdit) 
	{
		// TODO Auto-generated constructor stub
		setName(name);
		setPage(page);
		setCanEdit(canEdit);
	}
	
	
	public boolean isCanEdit() {
		return canEdit;
	}


	public void setCanEdit(boolean canEdit) {
		this.canEdit = canEdit;
	}


	public String getPage() {
		return page;
	}

	public void setPage(String page) {
		this.page = page;
	}

	public String getName() {
		return name;
	}

	public void setName(String name) {
		this.name = name;
	}
	
	public String getValue(String propName)
	{
		if(propName.equals("name"))
		{
			return name;
		}
		if(propName.equals("page"))
		{
			return page;
		}
		return "";
	}
	

}
