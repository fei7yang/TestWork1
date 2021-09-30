package com.dfl.report.handlers;



import java.awt.Frame;
import java.io.InputStream;

import com.dfl.report.util.ReportViwePanel;
import com.teamcenter.rac.aif.AbstractAIFUIApplication;
import com.teamcenter.rac.pse.AbstractPSEApplication;
import com.teamcenter.rac.util.MessageBox;


public class RobotAndWeldingExportAction implements Runnable {

	private AbstractAIFUIApplication app;
	private InputStream inputStream;
	//private AbstractPSEApplication pse;
	public RobotAndWeldingExportAction(AbstractAIFUIApplication app, Frame viewPanel, String string, InputStream inputStream) {
		// TODO Auto-generated constructor stub
		super();
		this.app=app;
		this.inputStream = inputStream;
	}

	@Override
	public void run() {
		// TODO Auto-generated method stub					
		RobotAndWeldingExportOp dlg = new RobotAndWeldingExportOp(app,inputStream);
		
	}

}
