package com.dfl.report.mfcadd;

import com.teamcenter.rac.kernel.TCComponent;
import com.teamcenter.rac.kernel.TCComponentBOMLine;

public class StdPartsSummaryReportAction implements Runnable{
	TCComponentBOMLine bopLine;
	TCComponent target;
	public StdPartsSummaryReportAction(TCComponentBOMLine bop, TCComponent folder) {
		bopLine = bop;
		target = folder;
	}
	@Override
	public void run() {
		// TODO Auto-generated method stub
		new StdPartsSummaryReportOperation(bopLine, target);
	}

}

