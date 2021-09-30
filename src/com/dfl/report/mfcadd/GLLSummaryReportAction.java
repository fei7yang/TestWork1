package com.dfl.report.mfcadd;

import com.teamcenter.rac.kernel.TCComponent;
import com.teamcenter.rac.kernel.TCComponentBOMLine;

public class GLLSummaryReportAction implements Runnable{
	TCComponentBOMLine bopLine;
	TCComponent target;
	public GLLSummaryReportAction(TCComponentBOMLine bop, TCComponent folder) {
		bopLine = bop;
		target = folder;
	}
	@Override
	public void run() {
		// TODO Auto-generated method stub
		new GLLSummaryReportOperation(bopLine, target);
	}

}
