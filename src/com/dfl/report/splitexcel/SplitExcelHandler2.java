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
		// TODO 自动生成的方法存根
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
			MessageBox.post("请选择文件夹或工艺报表对象!", "工艺报表拆分", MessageBox.WARNING);
			return;
		}

		boolean ispre = getSpecialChar();
		if (!ispre) {
			MessageBox.post("错误：首选项DFL_EngineeringWorkListSplitSheetName未定义，请联系系统管理员！", "工艺报表拆分", MessageBox.WARNING);
			return;
		}

		// 获取选中的工程作业表对象
		// 对象类型：DFL9MEDocument
		// 属性：dfl9_process_type=H(焊装工艺) & dfl9_process_file_type=AB(工程作业程序AB表)

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
						// 如果是工艺文档，就添加到集合，如果是文件夹，继续往下遍历，其他情况不处理
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
			MessageBox.post("请选择文件夹或工艺报表对象!", "工艺报表拆分", MessageBox.WARNING);
			return;
		}

		// 判断是否存在D盘
		File rootfile = new File("D:\\");
		if (!rootfile.exists()) {
			MessageBox.post("请在电脑上创建D盘", "工艺报表拆分", MessageBox.WARNING);
			return;
		}
		// 下载VBS脚本
		File scriptFile = Util.getRCPPluginInsideFile("UpdateSplitExcel.vbs");
		if (scriptFile == null || !scriptFile.exists()) {
			MessageBox.post("下载UpdateSplitExcel.vbs脚本错误", "工程作业表拆分", MessageBox.WARNING);
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
		System.out.println("文件夹：" + dialog.folder);

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
						monitor.setTaskName("操作进行中");

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
						MessageBox.post(resultmessage, "信息提示", MessageBox.WARNING);
						return;
					} else {
						MessageBox.post("拆分完成", "信息提示", MessageBox.WARNING);
						return;
					}
				} catch (Exception e2) {
					e2.printStackTrace();
				}
			}
		});
	}

	// 查询需要替换的特殊字符
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
