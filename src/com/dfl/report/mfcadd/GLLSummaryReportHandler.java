package com.dfl.report.mfcadd;

import org.eclipse.core.commands.AbstractHandler;
import org.eclipse.core.commands.ExecutionEvent;
import org.eclipse.core.commands.ExecutionException;
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

public class GLLSummaryReportHandler extends AbstractHandler {
	private AbstractAIFUIApplication app;
	private Shell shell;
	private TCSession session;
	private TCComponentFolder rootFolder;
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
		String inputStream = TemplateUtil.getTemplateFile("DFL_Template_GLLStatistics");
		if (inputStream == null) {
			MFCUtility.errorMassges("����û���ҵ�GLLͳ����Ϣ���ܱ��ģ�壬����ϵϵͳ����Ա����TC�����ģ�壨����Ϊ��DFL_Template_GLLStatistics��");
			//viewPanel.addInfomation("����û���ҵ�GLLͳ����Ϣ���ܱ��ģ�壬������TC�����ģ��(����Ϊ��DFL_Template_GLLStatistics)\n", 100,100);
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
		GLLSummaryReportAction action = new GLLSummaryReportAction(this.bopLine, savefolder);
		new Thread(action).start();
	}
}
