package com.dfl.report.mfcadd;

import com.teamcenter.rac.kernel.TCComponent;
import com.teamcenter.rac.kernel.TCComponentBOMLine;
import com.teamcenter.rac.kernel.TCComponentBOMWindow;

public class VehiclePSummaryReportAction implements Runnable{
	TCComponentBOMLine bopLine;
	TCComponent target;
	TCComponentBOMLine ebomLine;
	TCComponentBOMWindow bomwindow = null;
	public VehiclePSummaryReportAction(TCComponentBOMLine bop, TCComponentBOMLine ebom, TCComponentBOMWindow window, TCComponent folder) {
		bopLine = bop;
		target = folder;
		ebomLine = ebom;
		bomwindow = window;
	}
	@Override
	public void run() {
		// TODO Auto-generated method stub
		new VehiclePSummaryReportOperation(bopLine,ebomLine, bomwindow, target);
	}

}
