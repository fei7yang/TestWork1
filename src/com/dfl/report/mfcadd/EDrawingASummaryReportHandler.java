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
			MessageBox.post("������ֻ��ѡ��һ����װ�������հ汾����", "����", MessageBox.INFORMATION);
			return null;
		}
		if(aifComponents[0] instanceof TCComponentBOMLine) {
			bopLine = (TCComponentBOMLine)aifComponents[0];
			try {
				if (!bopLine.getItemRevision().isTypeOf("B8_BIWPlantBOPRevision")) {
					MessageBox.post("��ѡ�����д��ڲ��Ǻ�װ�������հ汾����", "��ܰ��ʾ", MessageBox.INFORMATION);
					return null;
				}
			} catch (TCException e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			}
		}else {
			MessageBox.post("��ѡ�����д��ڲ���BOMLine����", "��ʾ", MessageBox.INFORMATION);
			return null;
		}	
		try {
			if(bopLine != null && !bopLine.getItemRevision().okToModify()) {
				Util.callByPass(session, true);
//				MFCUtility.errorMassges("����ѡ��װ�������հ汾����û��дȨ�ޣ�");
//				return null;
				
			}
		} catch (Exception e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		String[] DFL_Project_VehicleNo = session.getPreferenceService().getStringValues("DFL_Project_VehicleNo");
		if(DFL_Project_VehicleNo == null || DFL_Project_VehicleNo.length == 0) {
			MFCUtility.errorMassges("��ѡ��δ�����δ���ã�DFL_Project_VehicleNo������ϵϵͳ����Ա��");
			return null;
		}
		String inputStream = TemplateUtil.getTemplateFile("DFL_Template_ManagementOfA");
		if (inputStream == null) {
			MFCUtility.errorMassges("����û���ҵ�������ͼA���ģ�壬����ϵϵͳ����Ա����TC�����ģ��(����Ϊ��DFL_Template_ManagementOfA)");
			return null;
		}
		String inputStreamS2 = TemplateUtil.getTemplateFile("DFL_Template_ManagementOfASheet2");
		if(inputStreamS2 == null || inputStreamS2.length() == 0) {
			MFCUtility.errorMassges("����û���ҵ�������ͼA��·��ͼ��ģ�壬����ϵϵͳ����Ա����TC�����ģ��(����Ϊ��DFL_Template_ManagementOfASheet2)");
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
