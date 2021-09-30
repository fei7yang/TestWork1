package com.dfl.report.defects;

import com.dfl.report.dialog.SelectionStageDialog;
import com.teamcenter.rac.aif.AbstractAIFUIApplication;
import com.teamcenter.rac.aif.kernel.InterfaceAIFComponent;
import com.teamcenter.rac.kernel.TCComponent;
import com.teamcenter.rac.kernel.TCSession;

public class DefectsListAction implements Runnable{
	
	private AbstractAIFUIApplication app;
	private TCComponent folder;
	private InterfaceAIFComponent[] aifComponents;
	private TCSession session;
	
	public DefectsListAction(AbstractAIFUIApplication app, Object object, TCComponent savefolder, InterfaceAIFComponent[] aifComponents, TCSession session) {
		// TODO Auto-generated constructor stub
		this.app=app;
		this.folder=savefolder;
		this.aifComponents = aifComponents;
		this.session = session;
	}
	@Override
	public void run() {
		// TODO Auto-generated method stub		
		DefectsListDialog dialog = new DefectsListDialog(app,folder,aifComponents,session);
		dialog.show();
	}

}
