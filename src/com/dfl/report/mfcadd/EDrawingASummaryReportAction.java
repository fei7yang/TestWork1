package com.dfl.report.mfcadd;

import com.teamcenter.rac.kernel.TCComponent;
import com.teamcenter.rac.kernel.TCComponentBOMLine;

public class EDrawingASummaryReportAction implements Runnable{
	TCComponentBOMLine bopLine;
	TCComponent target;
	String version = "";
	public EDrawingASummaryReportAction(TCComponentBOMLine bop, TCComponent folder, String ver) {
		bopLine = bop;
		target = folder;
		version = ver;
	}
	@Override
	public void run() {
		// TODO Auto-generated method stub
		new EDrawingASummaryReportOperation(bopLine, target, version);
	}

}
