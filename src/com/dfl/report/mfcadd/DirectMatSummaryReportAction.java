package com.dfl.report.mfcadd;

import com.teamcenter.rac.kernel.TCComponent;
import com.teamcenter.rac.kernel.TCComponentBOMLine;

public class DirectMatSummaryReportAction implements Runnable{
	TCComponentBOMLine bopLine;
	TCComponentBOMLine[]	 virtualLines;
	TCComponent target;
	public DirectMatSummaryReportAction(TCComponentBOMLine bop, TCComponentBOMLine[] lines, TCComponent folder) {
		bopLine = bop;
		virtualLines = lines;
		target = folder;
	}
	@Override
	public void run() {
		// TODO Auto-generated method stub
		new DirectMatSummaryReportOperation(bopLine, virtualLines, target);
	}

}
