package com.dfl.report.home;

import com.teamcenter.rac.kernel.TCComponent;
import com.teamcenter.rac.kernel.TCException;



public class TreeNode  {

	private String name;
	private TCComponent folder;

	public TreeNode(TCComponent folder) {
		// TODO Auto-generated constructor stub
		setFolder(folder);
		try {
			setName(folder.getProperty("object_name"));
		} catch (TCException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
	}

	public String getName() {
		return name;
	}

	public void setName(String name) {
		this.name = name;
	}

	public TCComponent getFolder() {
		return folder;
	}

	public void setFolder(TCComponent folder) {
		this.folder = folder;
	}

}
