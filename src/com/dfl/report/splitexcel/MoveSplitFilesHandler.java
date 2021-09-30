package com.dfl.report.splitexcel;

import java.io.File;
import java.util.ArrayList;
import java.util.List;

import org.eclipse.core.commands.AbstractHandler;
import org.eclipse.core.commands.ExecutionEvent;
import org.eclipse.core.commands.ExecutionException;
import org.eclipse.core.runtime.IProgressMonitor;
import org.eclipse.jface.dialogs.ProgressMonitorDialog;
import org.eclipse.jface.operation.IRunnableWithProgress;
import org.eclipse.swt.widgets.Display;
import org.eclipse.swt.widgets.Shell;

import com.dfl.report.home.OpenHomeDialog;
import com.dfl.report.util.Util;
import com.teamcenter.rac.aif.AIFDesktop;
import com.teamcenter.rac.aif.AbstractAIFUIApplication;
import com.teamcenter.rac.aif.kernel.InterfaceAIFComponent;
import com.teamcenter.rac.kernel.TCComponent;
import com.teamcenter.rac.kernel.TCComponentFolder;
import com.teamcenter.rac.kernel.TCComponentItem;
import com.teamcenter.rac.kernel.TCComponentItemRevision;
import com.teamcenter.rac.kernel.TCException;
import com.teamcenter.rac.kernel.TCSession;
import com.teamcenter.rac.util.MessageBox;

public class MoveSplitFilesHandler extends AbstractHandler{

	@Override
	public Object execute(ExecutionEvent arg0) throws ExecutionException {
		// TODO Auto-generated method stub
		System.out.println("--------------MoveSplitFilesHandler---------------");
		Thread thread = new Thread() {
			public void run() {
				execute();
			}
		};
		thread.start();
		return null;
	}
	private AbstractAIFUIApplication application;
	private TCSession session;
	private TCComponentFolder rootFolder;
	private ProgressMonitorDialog monitorDialog;
	private MoveSplitFilesOperation MoveSplitFilesOperation;
	private Shell shell;
	private TCComponent savefolder;
	private List<TCComponentFolder> splitlist ;
	
	protected void execute() {
		// TODO Auto-generated method stub
		application = AIFDesktop.getActiveDesktop().getCurrentApplication();
		session = (TCSession) application.getSession();
		shell = application.getDesktop().getShell();

		InterfaceAIFComponent[] aifComponents = application.getTargetComponents();

		if (aifComponents == null || aifComponents.length <= 0) {
			MessageBox.post("请选择文件夹对象!", "移动拆分报表文件", MessageBox.WARNING);
			return;
		}		
		TCComponentFolder folder;
		splitlist = new ArrayList<TCComponentFolder>();
		for (int i = 0; i < aifComponents.length; i++) {
			if (aifComponents[i] instanceof TCComponentFolder) {
				folder = (TCComponentFolder) aifComponents[i];
				splitlist.add(folder);
			}
		}
		if(splitlist.size()<1) {
			MessageBox.post("请选择文件夹对象!", "移动拆分报表文件", MessageBox.WARNING);
			return;
		}
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

						MoveSplitFilesOperation = new MoveSplitFilesOperation(session, splitlist, savefolder);

						try {
							MoveSplitFilesOperation.executeOperation();
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
					String resultmessage = MoveSplitFilesOperation.getResultMessage();
					if(resultmessage!=null && !resultmessage.isEmpty()) {
						MessageBox.post(resultmessage, "信息提示", MessageBox.WARNING);
						return;
					}else {
						MessageBox.post("文件移动完成", "信息提示", MessageBox.WARNING);
						return;
					}
				} catch (Exception e2) {
					e2.printStackTrace();
				}
			}
		});
	}

}
