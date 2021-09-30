package com.dfl.report.straightmaterial;

import com.teamcenter.rac.aif.AbstractAIFUIApplication;

public class StraightMaterialCQAction implements Runnable {

	private AbstractAIFUIApplication app;

	public StraightMaterialCQAction(AbstractAIFUIApplication app, Object object, String string) {
		// TODO Auto-generated constructor stub
		super();
		this.app=app;
	}

	@Override
	public void run() {
		// TODO Auto-generated method stub
		StraightMaterialCQOp op = new StraightMaterialCQOp(app);
	}

}
