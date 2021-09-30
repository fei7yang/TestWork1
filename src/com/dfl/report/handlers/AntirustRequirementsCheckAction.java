package com.dfl.report.handlers;

import java.util.ArrayList;

import com.dfl.report.dialog.SelectionStageDialog;
import com.teamcenter.rac.aif.AbstractAIFUIApplication;
import com.teamcenter.rac.aif.kernel.InterfaceAIFComponent;
import com.teamcenter.rac.kernel.TCComponent;
import com.teamcenter.rac.kernel.TCSession;

public class AntirustRequirementsCheckAction implements Runnable {

	private AbstractAIFUIApplication app;
	private TCComponent folder;
	private InterfaceAIFComponent[] ifc;
	private TCSession session;
	private ArrayList rule;

	public AntirustRequirementsCheckAction(AbstractAIFUIApplication app, Object object, TCComponent savefolder, InterfaceAIFComponent[] ifc, TCSession session, ArrayList rule) {
		// TODO Auto-generated constructor stub
		super();
		this.app=app;
		this.folder = savefolder;
		this.ifc = ifc;
		this.session = session;
		this.rule = rule;
	}

	@SuppressWarnings("deprecation")
	@Override
	public void run() {
		// TODO Auto-generated method stub		
		SelectionStageDialog dialog = new SelectionStageDialog(app,folder,ifc,session,rule);
		dialog.show();
	}

}
