package com.dfl.report.mfcadd;

import com.teamcenter.rac.kernel.TCComponent;
import com.teamcenter.rac.kernel.TCComponentBOMLine;

public class DATUMSummaryReportAction implements Runnable{
	TCComponentBOMLine bopLine;
	TCComponent target;
	public DATUMSummaryReportAction(TCComponentBOMLine bop, TCComponent folder) {
		bopLine = bop;
		target = folder;
	}
	@Override
	public void run() {
		// TODO Auto-generated method stub
		new DATUMSummaryReportOperation(bopLine, target);
	}

}
