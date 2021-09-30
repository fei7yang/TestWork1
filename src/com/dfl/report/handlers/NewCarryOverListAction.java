package com.dfl.report.handlers;

import com.teamcenter.rac.aif.AbstractAIFUIApplication;
import com.teamcenter.rac.aif.kernel.InterfaceAIFComponent;
import com.teamcenter.rac.kernel.TCComponent;
import com.teamcenter.rac.kernel.TCSession;

public class NewCarryOverListAction implements Runnable{
	private AbstractAIFUIApplication app;
	private TCComponent folder;
	private InterfaceAIFComponent[] aifComponents;
	private TCSession session;
	
	public NewCarryOverListAction(AbstractAIFUIApplication app, Object object, TCComponent savefolder, InterfaceAIFComponent[] aifComponents, TCSession session) {
		// TODO Auto-generated constructor stub
		this.app=app;
		this.folder=savefolder;
		this.aifComponents = aifComponents;
		this.session =session;
	}

	@Override
	public void run() {
		// TODO Auto-generated method stub
		NewCarryOverListOp operatioon= new NewCarryOverListOp(app,folder,aifComponents,session);
	}

}