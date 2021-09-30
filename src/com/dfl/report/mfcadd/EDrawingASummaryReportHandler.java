package com.dfl.report.mfcadd;

import org.eclipse.core.commands.AbstractHandler;
import org.eclipse.core.commands.ExecutionEvent;
import org.eclipse.core.commands.ExecutionException;
import org.eclipse.swt.SWT;
import org.eclipse.swt.widgets.Display;
import org.eclipse.swt.widgets.Shell;

import com.dfl.report.home.OpenHomeDialog;
import com.dfl.report.util.Util;
import com.teamcenter.rac.aif.AIFDesktop;
import com.teamcenter.rac.aif.AbstractAIFUIApplication;
import com.teamcenter.rac.aif.kernel.InterfaceAIFComponent;
import com.teamcenter.rac.aifrcp.AIFUtility;
import com.teamcenter.rac.kernel.TCComponent;
import com.teamcenter.rac.kernel.TCComponentBOMLine;
import com.teamcenter.rac.kernel.TCComponentFolder;
import com.teamcenter.rac.kernel.TCException;
import com.teamcenter.rac.kernel.TCSession;
import com.teamcenter.rac.util.MessageBox;

public class EDrawingASummaryReportHandler extends AbstractHandler {
	private AbstractAIFUIApplication app;
	private Shell shell;
	private TCSession session;
	private TCComponent savefolder;
	TCComponentBOMLine bopLine = null;
	@Override
	public Object execute(ExecutionEvent arg0) throws ExecutionException {
		// TODO Auto-generated method stub
		app = AIFUtility.getCurrentApplication();
		session = (TCSession)AIFUtility.getDefaultSession();
		InterfaceAIFComponent[] aifComponents = app.getTargetComponents();
		if (aifComponents == null || aifComponents.length != 1) {
			MessageBox.post("必须且只能选择一个焊装工厂工艺版本对象！", "错误", MessageBox.INFORMATION);
			return null;
		}
		if(aifComponents[0] instanceof TCComponentBOMLine) {
			bopLine = (TCComponentBOMLine)aifComponents[0];
			try {
				if (!bopLine.getItemRevision().isTypeOf("B8_BIWPlantBOPRevision")) {
					MessageBox.post("所选对象中存在不是焊装工厂工艺版本对象！", "温馨提示", MessageBox.INFORMATION);
					return null;
				}
			} catch (TCException e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			}
		}else {
			MessageBox.post("所选对象中存在不是BOMLine对象！", "提示", MessageBox.INFORMATION);
			return null;
		}	
		try {
			if(bopLine != null && !bopLine.getItemRevision().okToModify()) {
				Util.callByPass(session, true);
//				MFCUtility.errorMassges("对所选焊装工厂工艺版本对象没有写权限！");
//				return null;
				
			}
		} catch (Exception e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		String[] DFL_Project_VehicleNo = session.getPreferenceService().getStringValues("DFL_Project_VehicleNo");
		if(DFL_Project_VehicleNo == null || DFL_Project_VehicleNo.length == 0) {
			MFCUtility.errorMassges("首选项未定义或未配置：DFL_Project_VehicleNo，请联系系统管理员！");
			return null;
		}
		String inputStream = TemplateUtil.getTemplateFile("DFL_Template_ManagementOfA");
		if (inputStream == null) {
			MFCUtility.errorMassges("错误：没有找到管理工程图A表的模板，请联系系统管理员先在TC中添加模板(名称为：DFL_Template_ManagementOfA)");
			return null;
		}
		String inputStreamS2 = TemplateUtil.getTemplateFile("DFL_Template_ManagementOfASheet2");
		if(inputStreamS2 == null || inputStreamS2.length() == 0) {
			MFCUtility.errorMassges("错误：没有找到管理工程图A表路径图的模板，请联系系统管理员先在TC中添加模板(名称为：DFL_Template_ManagementOfASheet2)");
			return null;
		}
		Thread thread = new Thread() {
			public void run() {
				execute();				
			}
		};
		thread.start();
		return null;
	}
	protected void execute() {
		// TODO Auto-generated method stub
		session = (TCSession) app.getSession();
		shell = AIFDesktop.getActiveDesktop().getShell();

		//rootFolder = (TCComponent) aifComponent;

		Display.getDefault().asyncExec(new Runnable(){
			@Override
			public void run() {
				openDialog();
			}
		});


	}
	protected void openDialog() {
		// TODO Auto-generated method stub
		VersionSelectionDialog dialog = new VersionSelectionDialog(new Shell(), SWT.SHELL_TRIM);
		dialog.open();
		String version = dialog.version;
		if(StringUtil.isEmpty(dialog.version)) {
			return;
		}
		EDrawingASummaryReportAction action = new EDrawingASummaryReportAction(this.bopLine, savefolder, version);
		new Thread(action).start();
	}

}
