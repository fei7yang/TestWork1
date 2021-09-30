package com.dfl.report.splitexcel;

import java.io.File;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

import org.eclipse.core.commands.AbstractHandler;
import org.eclipse.core.commands.ExecutionEvent;
import org.eclipse.core.commands.ExecutionException;
import org.eclipse.core.runtime.IProgressMonitor;
import org.eclipse.jface.dialogs.ProgressMonitorDialog;
import org.eclipse.jface.operation.IRunnableWithProgress;
import org.eclipse.swt.SWT;
import org.eclipse.swt.widgets.Display;
import org.eclipse.swt.widgets.Shell;

import com.dfl.report.DotStatistics.DotStatisticsOp;
import com.dfl.report.home.OpenHomeDialog;
import com.dfl.report.util.Util;
import com.teamcenter.rac.aif.AIFDesktop;
import com.teamcenter.rac.aif.AbstractAIFUIApplication;
import com.teamcenter.rac.aif.kernel.AIFComponentContext;
import com.teamcenter.rac.aif.kernel.InterfaceAIFComponent;
import com.teamcenter.rac.aifrcp.AIFUtility;
import com.teamcenter.rac.kernel.TCComponent;
import com.teamcenter.rac.kernel.TCComponentFolder;
import com.teamcenter.rac.kernel.TCComponentItem;
import com.teamcenter.rac.kernel.TCComponentItemRevision;
import com.teamcenter.rac.kernel.TCException;
import com.teamcenter.rac.kernel.TCPreferenceService;
import com.teamcenter.rac.kernel.TCSession;
import com.teamcenter.rac.util.MessageBox;
import com.teamcenter.soa.client.model.ErrorStack;

public class SplitExcelHandler2 extends AbstractHandler {

	private AbstractAIFUIApplication application;
	private TCSession session;
	private TCComponentFolder rootFolder;
	private TCComponent savefolder;

	@Override
	public Object execute(ExecutionEvent event) throws ExecutionException {
		// TODO �Զ����ɵķ������
		System.out.println("--------------SplitExcelHandler---------------");
		Thread thread = new Thread() {
			public void run() {
				try {
					execute();
				} catch (TCException e) {
					// TODO Auto-generated catch block
					e.printStackTrace();
				}
			}
		};
		thread.start();
		return null;
	}

	private ProgressMonitorDialog monitorDialog;
	private SplitExcelOperation2 splitExcelOperation2;
	private Shell shell;
	private ArrayList<TCComponentItem> documentList;
	private String vbsFilePath;
	private String alterordercode;

	protected void execute() throws TCException {
		// TODO Auto-generated method stub
		application = AIFDesktop.getActiveDesktop().getCurrentApplication();
		session = (TCSession) application.getSession();
		shell = application.getDesktop().getShell();

		InterfaceAIFComponent[] aifComponents = application.getTargetComponents();

		if (aifComponents == null || aifComponents.length <= 0) {
			MessageBox.post("��ѡ���ļ��л��ձ������!", "���ձ�����", MessageBox.WARNING);
			return;
		}

		boolean ispre = getSpecialChar();
		if (!ispre) {
			MessageBox.post("������ѡ��DFL_EngineeringWorkListSplitSheetNameδ���壬����ϵϵͳ����Ա��", "���ձ�����", MessageBox.WARNING);
			return;
		}

		// ��ȡѡ�еĹ�����ҵ�����
		// �������ͣ�DFL9MEDocument
		// ���ԣ�dfl9_process_type=H(��װ����) & dfl9_process_file_type=AB(������ҵ����AB��)

		documentList = new ArrayList<TCComponentItem>();
		TCComponentItem item;
		String type;
		TCComponentItemRevision rev;
		// dfl9_process_type
		String process_type;
		String process_file_type;

		for (int i = 0; i < aifComponents.length; i++) {
			if (aifComponents[i] instanceof TCComponentItem) {
				item = (TCComponentItem) aifComponents[i];
				type = item.getType();
				if ("DFL9MEDocument".equals(type)) {
					if (!documentList.contains(item)) {
						documentList.add(item);
					}
				}
			}
			if (aifComponents[i] instanceof TCComponentFolder) {
				TCComponent folder = (TCComponent) aifComponents[i];
				AIFComponentContext[] contexts = folder.getRelated("contents");
				for (AIFComponentContext aif : contexts) {
					TCComponent tcc = (TCComponent) aif.getComponent();
					if (tcc instanceof TCComponentItem) {
						String objecttype = tcc.getType();
						// ����ǹ����ĵ�������ӵ����ϣ�������ļ��У��������±������������������
						if (objecttype.equals("DFL9MEDocument")) {
							item = (TCComponentItem) tcc;
							if (!documentList.contains(item)) {
								documentList.add(item);
							}
						}
					}
				}
			}
		}

		if (documentList.isEmpty()) {
			MessageBox.post("��ѡ���ļ��л��ձ������!", "���ձ�����", MessageBox.WARNING);
			return;
		}

		// �ж��Ƿ����D��
		File rootfile = new File("D:\\");
		if (!rootfile.exists()) {
			MessageBox.post("���ڵ����ϴ���D��", "���ձ�����", MessageBox.WARNING);
			return;
		}
		// ����VBS�ű�
		File scriptFile = Util.getRCPPluginInsideFile("UpdateSplitExcel.vbs");
		if (scriptFile == null || !scriptFile.exists()) {
			MessageBox.post("����UpdateSplitExcel.vbs�ű�����", "������ҵ����", MessageBox.WARNING);
			return;
		}

		vbsFilePath = scriptFile.getAbsolutePath();

		shell = AIFDesktop.getActiveDesktop().getShell();

		try {
			rootFolder = session.getUser().getHomeFolder();
		} catch (TCException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		Display.getDefault().asyncExec(new Runnable() {
			@Override
			public void run() {
				openDialog();
			}
		});

	}

	protected void openDialog() {
		// TODO Auto-generated method stub
		splitDialog splitdialog = new splitDialog(shell, SWT.SHELL_TRIM);
		splitdialog.open();
		alterordercode = splitdialog.altercode;

		if (alterordercode == null) {
			return;
		}

		OpenHomeDialog dialog = new OpenHomeDialog(shell, rootFolder, session);
		dialog.open();

		savefolder = dialog.folder;
		System.out.println("�ļ��У�" + dialog.folder);

		if (dialog.flag) {
			return;
		}

		if (savefolder == null) {
			return;
		}

		Display.getDefault().asyncExec(new Runnable() {
			@Override
			public void run() {
				IRunnableWithProgress runnable = new IRunnableWithProgress() {
					public void run(IProgressMonitor monitor) {
						// monitor.beginTask("begin" + "...... ",10);
						monitor.beginTask("......", IProgressMonitor.UNKNOWN);
						monitor.setTaskName("����������");

						splitExcelOperation2 = new SplitExcelOperation2(session, documentList, vbsFilePath, savefolder,
								alterordercode);

						try {
							splitExcelOperation2.executeOperation();
						} catch (Exception e) {
							// TODO Auto-generated catch block
							monitor.done();
							e.printStackTrace();
							MessageBox.post(e);
						}
						System.out.println("monitor done");
						monitor.done();
					}
				};

				try {
					monitorDialog = new ProgressMonitorDialog(new Shell());
					monitorDialog.run(true, true, runnable);
					System.out.println("monitorDialog finish");
					String resultmessage = splitExcelOperation2.getResultMessage();
					if (resultmessage != null && !resultmessage.isEmpty()) {
						MessageBox.post(resultmessage, "��Ϣ��ʾ", MessageBox.WARNING);
						return;
					} else {
						MessageBox.post("������", "��Ϣ��ʾ", MessageBox.WARNING);
						return;
					}
				} catch (Exception e2) {
					e2.printStackTrace();
				}
			}
		});
	}

	// ��ѯ��Ҫ�滻�������ַ�
	private boolean getSpecialChar() {
		try {
			AbstractAIFUIApplication app = AIFUtility.getCurrentApplication();
			TCSession session = (TCSession) app.getSession();
			TCPreferenceService preferenceService = session.getPreferenceService();

			String str = preferenceService.getPreferenceDescription("DFL_EngineeringWorkListSplitSheetName");
			if (str != null) {
				return true;
			}
		} catch (Exception e) {
			e.printStackTrace();
		}
		return false;
	}
}
