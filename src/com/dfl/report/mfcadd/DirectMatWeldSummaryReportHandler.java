package com.dfl.report.mfcadd;

import java.io.File;
import java.util.ArrayList;

import org.eclipse.core.commands.AbstractHandler;
import org.eclipse.core.commands.ExecutionEvent;
import org.eclipse.core.commands.ExecutionException;
import org.eclipse.swt.SWT;
import org.eclipse.swt.widgets.Display;
import org.eclipse.swt.widgets.Shell;

import com.dfl.report.home.OpenHomeDialog;
import com.teamcenter.rac.aif.AIFDesktop;
import com.teamcenter.rac.aif.AbstractAIFUIApplication;
import com.teamcenter.rac.aif.kernel.AIFComponentContext;
import com.teamcenter.rac.aif.kernel.InterfaceAIFComponent;
import com.teamcenter.rac.aifrcp.AIFUtility;
import com.teamcenter.rac.kernel.TCComponent;
import com.teamcenter.rac.kernel.TCComponentBOMLine;
import com.teamcenter.rac.kernel.TCComponentFolder;
import com.teamcenter.rac.kernel.TCComponentQuery;
import com.teamcenter.rac.kernel.TCComponentQueryType;
import com.teamcenter.rac.kernel.TCException;
import com.teamcenter.rac.kernel.TCSession;
import com.teamcenter.rac.util.MessageBox;

public class DirectMatWeldSummaryReportHandler extends AbstractHandler {
	private AbstractAIFUIApplication app;
	private Shell shell;
	private TCSession session;
	private TCComponentFolder rootFolder;
	private TCComponent savefolder;
	TCComponentBOMLine bopLine = null;
	private TCComponentBOMLine[] selectLines;
	@Override
	public Object execute(ExecutionEvent arg0) throws ExecutionException {
		// TODO Auto-generated method stub
		app = AIFUtility.getCurrentApplication();
		selectLines = null;
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
		String[] DFL_Project_VehicleNo = session.getPreferenceService().getStringValues("DFL_Project_VehicleNo");
		if(DFL_Project_VehicleNo == null || DFL_Project_VehicleNo.length == 0) {
			MFCUtility.errorMassges("��ѡ��δ�����δ���ã�DFL_Project_VehicleNo������ϵϵͳ����Ա��");
			return null;
		}
		String prefValue = session.getPreferenceService().getStringValue("DFL9_DirectMate_CountRule");
		if(prefValue == null || prefValue.length() == 0) {
			MFCUtility.errorMassges("��ѡ��δ�����δ���ã�DFL9_DirectMate_CountRule������ϵϵͳ����Ա��");
			return null;
		}
		String countRule = TemplateUtil.getTemplateFile(prefValue);
		boolean downerror = true;
		if (countRule == null) {
			downerror = false;
		}else {
			File file = new File(countRule);
			if(file.exists()) {
				file.delete();
			}else {
				downerror  = false;
			}
		}
		if(!downerror) {
			MFCUtility.errorMassges("û���ҵ�ֱ�ļ�������Excel�ļ�������ϵ����Ա��TC����ӣ�����Ϊ��" + prefValue + "��MSExcelX���ݼ���");
			return null;
		}
		String inputStream = TemplateUtil.getTemplateFile("DFL_Template_HZDirectMetaList");
		if (inputStream == null) {
			downerror = false;
		}else {
			File file = new File(inputStream);
			if(file.exists()) {
				file.delete();
			}else {
				downerror = false;
			}
		}
		if(!downerror) {
			MFCUtility.errorMassges("����û���ҵ�ֱ���嵥����������ģ�壬������TC�����ģ��(����Ϊ��DFL_Template_HZDirectMetaList)������ϵϵͳ����Ա��" );
			return null;
		}
		try {
			TCComponentQueryType queryType = (TCComponentQueryType) session.getTypeComponent("ImanQuery");
			TCComponentQuery query = (TCComponentQuery) queryType.find("__DFL_Find_Object_by_Name");
			if (query == null) {
				MFCUtility.errorMassges("��ѯδ���壺__DFL_Find_Object_by_Name������ϵϵͳ����Ա��" );
				return null;
			}
		}catch(Exception e) {
			e.printStackTrace();
		}
		Thread thread = new Thread() {
			public void run() {
				execute();				
			}
		};
		thread.start();
		return null;
	}
	private boolean isVirtualLine(TCComponentBOMLine lineLine) {
		boolean isVir = true;
		try {
			AIFComponentContext[] children = lineLine.getChildren();
			int i = 0;
			int count = children.length;
			for(i = 0; i < count; i ++) {
				TCComponentBOMLine cline = (TCComponentBOMLine)children[i].getComponent();
				if(cline.getItem().getType().equals("B8_BIWMEProcStat")) {
					isVir = false;
					break;
				}
			}
		}catch(Exception e) {
			e.printStackTrace();
		}
		return isVir;
	}
	protected void execute() {
		// TODO Auto-generated method stub
		session = (TCSession) app.getSession();
		shell = AIFDesktop.getActiveDesktop().getShell();

		//InterfaceAIFComponent aifComponent = app.getTargetComponent();
		
		try {
			rootFolder = session.getUser().getHomeFolder();
		} catch (TCException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}

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
		OpenHomeDialog dialog = new OpenHomeDialog(shell, rootFolder,session);
		dialog.open();
		
		savefolder = dialog.folder;
		System.out.println("�ļ��У�"+dialog.folder);
		
		if(dialog.flag) {
			return ;
		}
		
		if(savefolder==null ) {
			return ;
		}
		VersionSelectionDialog dialog2 = new VersionSelectionDialog(new Shell(), SWT.SHELL_TRIM);
		dialog2.open();
		String version = dialog2.version;
		if(StringUtil.isEmpty(dialog2.version)) {
			//MFCUtility.errorMassges("����ѡ��һ���Σ�");
			return;
		}
		DirectMatWeldSummaryReportAction action = new DirectMatWeldSummaryReportAction(this.bopLine, selectLines, savefolder, version);
		new Thread(action).start();
	}
}
