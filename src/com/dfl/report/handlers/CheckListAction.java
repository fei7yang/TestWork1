package com.dfl.report.handlers;

import java.awt.Frame;
import java.io.InputStream;
import java.util.ArrayList;

import com.dfl.report.dialog.SelectionStageDialog;
import com.teamcenter.rac.aif.AbstractAIFUIApplication;
import com.teamcenter.rac.aif.kernel.InterfaceAIFComponent;
import com.teamcenter.rac.kernel.TCComponent;
import com.teamcenter.rac.kernel.TCSession;

public class CheckListAction implements Runnable {

	private AbstractAIFUIApplication app;
	private TCComponent folder;
	private InterfaceAIFComponent[] aifComponents;
	private TCSession session;
	private ArrayList rule;
	public CheckListAction(AbstractAIFUIApplication app, Frame viewPanel, TCComponent savefolder, InterfaceAIFComponent[] aifComponents, TCSession session, ArrayList rule) {
		// TODO Auto-generated constructor stub
		super();
		this.app=app;
		this.folder=savefolder;
		this.aifComponents = aifComponents;
		this.session = session;
		this.rule = rule;
	}

	@Override
	public void run() {
		// TODO Auto-generated method stub					
		CheckListDialog dialog = new CheckListDialog(app,folder,aifComponents,session,rule);
		dialog.show();
		
	}

}