package com.dfl.report.handlers;

import java.io.InputStream;

import com.teamcenter.rac.aif.AbstractAIFDialog;
import com.teamcenter.rac.aif.AbstractAIFUIApplication;
import com.teamcenter.rac.aif.kernel.InterfaceAIFComponent;
import com.teamcenter.rac.kernel.TCComponent;
import com.teamcenter.rac.kernel.TCSession;

public class StraightforwardlistAction  implements Runnable {

	private AbstractAIFUIApplication app;
	private TCComponent folder;
	private InterfaceAIFComponent[] ifc;
	private TCSession session;
	private InputStream inputStream;

	public StraightforwardlistAction() {
		// TODO Auto-generated constructor stub
	}

	public StraightforwardlistAction(AbstractAIFUIApplication app, Object object, TCComponent savefolder, InterfaceAIFComponent[] ifc, TCSession session,InputStream inputStream) {
		// TODO Auto-generated constructor stub
		super();
		this.app=app;
		this.folder = savefolder;
		this.ifc= ifc;
		this.session = session;
		this.inputStream = inputStream;
	}

	@Override
	public void run() {
		// TODO Auto-generated method stub
		StraightforwardlistOp operation = new StraightforwardlistOp(app,folder,ifc,session,inputStream);
	}

}
