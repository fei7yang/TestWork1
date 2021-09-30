package com.dfl.report.mfcadd;

import com.teamcenter.rac.kernel.TCComponent;
import com.teamcenter.rac.kernel.TCComponentBOMLine;

public class DirectMatWeldSummaryReportAction implements Runnable{
	TCComponentBOMLine bopLine;
	TCComponentBOMLine[]	 virtualLines;
	TCComponent target;
	String version = "";
	public DirectMatWeldSummaryReportAction(TCComponentBOMLine bop, TCComponentBOMLine[] lines, 
			TCComponent folder, String ver) {
		bopLine = bop;
		virtualLines = lines;
		target = folder;
		version = ver;
	}
	@Override
	public void run() {
		// TODO Auto-generated method stub
		new DirectMatWeldSummaryReportOperation(bopLine, virtualLines, target, version);
	}

}

