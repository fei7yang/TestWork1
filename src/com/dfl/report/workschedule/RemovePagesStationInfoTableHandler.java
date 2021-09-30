package com.dfl.report.workschedule;

import java.io.InputStream;
import java.util.ArrayList;
import java.util.List;

import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.eclipse.core.commands.AbstractHandler;
import org.eclipse.core.commands.ExecutionEvent;
import org.eclipse.core.commands.ExecutionException;
import org.eclipse.swt.SWT;
import org.eclipse.swt.widgets.Display;
import org.eclipse.swt.widgets.Shell;

import com.dfl.report.ExcelReader.baseinfoExcelReader;
import com.dfl.report.util.GenerateReportInfo;
import com.dfl.report.util.NewOutputDataToExcel;
import com.dfl.report.util.ReportUtils;
import com.dfl.report.util.Util;
import com.teamcenter.rac.aif.AIFDesktop;
import com.teamcenter.rac.aif.AbstractAIFUIApplication;
import com.teamcenter.rac.aif.kernel.InterfaceAIFComponent;
import com.teamcenter.rac.aifrcp.AIFUtility;
import com.teamcenter.rac.kernel.TCComponentBOMLine;
import com.teamcenter.rac.kernel.TCComponentItemRevision;
import com.teamcenter.rac.kernel.TCException;
import com.teamcenter.rac.kernel.TCSession;
import com.teamcenter.rac.util.MessageBox;

/*
 * 工程作业本减页操作
 */
public class RemovePagesStationInfoTableHandler extends AbstractHandler {
	private AbstractAIFUIApplication app;
	private Shell shell;
	private List shList;
	private TCSession session;
	private XSSFWorkbook book;
	private TCComponentBOMLine topbomline;
	private GenerateReportInfo info;
	private List sheetlist;
	@Override
	public Object execute(ExecutionEvent arg0) throws ExecutionException {
		// TODO Auto-generated method stub
		app = AIFUtility.getCurrentApplication();
		session = (TCSession) app.getSession();
		InterfaceAIFComponent[] ifc = app.getTargetComponents();
		if (ifc.length < 1) {
			MessageBox.post("当前未选择操作对象，请先选择！", "温馨提示", MessageBox.INFORMATION);
			return null;
		}
		if (ifc.length > 1) {
			MessageBox.post("请选择单一焊装工位工艺对象！", "温馨提示", MessageBox.INFORMATION);
			return null;
		}
		if (ifc[0] instanceof TCComponentBOMLine) {

		} else {
			MessageBox.post("请选择焊装工位工艺对象！", "提示", MessageBox.INFORMATION);
			return null;
		}
		topbomline = (TCComponentBOMLine) ifc[0];

		try {
			System.out.println(topbomline.getItemRevision().getType());
			if (!topbomline.getItemRevision().isTypeOf("B8_BIWMEProcStatRevision")) {
				MessageBox.post("请选择焊装工厂工位对象！", "温馨提示", MessageBox.INFORMATION);
				return null;
			}
		} catch (TCException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		
		TCComponentItemRevision toprev = null;
		try {
			toprev = topbomline.window().getTopBOMLine().getItemRevision();
		} catch (TCException e1) {
			// TODO Auto-generated catch block
			e1.printStackTrace();
		}
		
		// 获取已生成的报表
		info = new GenerateReportInfo();
		info.setExist(false);
		info.setIsgoon(true);
		info.setAction(""); //$NON-NLS-1$
		info.setMeDocument(null);
		info.setDFL9_process_type("H"); //$NON-NLS-1$
		info.setDFL9_process_file_type("AB"); // $NON-NLS-1$
		info.setmeDocumentName("");
		info.setFlag(true);
		info.setProject_ids(toprev);

		try {
			info = ReportUtils.beforeGenerateReportAction(topbomline.getItemRevision(), info);
		} catch (TCException e) {
			e.printStackTrace();
			// EclipseUtils.info("Error : " + e.getMessage()); //$NON-NLS-1$
			return null;
		}
		System.out.println("The action is completed before the report operation is generated.");

		if (!info.isIsgoon()) {
			return null;
		}
		TCComponentItemRevision docmentRev = info.getMeDocument();
		if(docmentRev ==null) {
			MessageBox.post("请确认已经生成工程作业表！", "温馨提示", MessageBox.INFORMATION);
			return null;
		}
		String procName = Util.getProperty(docmentRev, "object_name");
		InputStream inputStream = baseinfoExcelReader.getFileinbyreadExcel(docmentRev, "IMAN_specification", procName);

		if (inputStream == null) {
			MessageBox.post("请确认" + procName + "版本对象下，存在" + procName + "数据集！", "温馨提示", MessageBox.INFORMATION);
			return null;
		}
		book = NewOutputDataToExcel.creatXSSFWorkbook(inputStream);
		shList = new ArrayList();
		for (int i = 0; i < book.getNumberOfSheets(); i++) {
			XSSFSheet sheet = book.getSheetAt(i);
			String sheetname = sheet.getSheetName();
			shList.add(sheetname);
		}
		Display.getDefault().asyncExec(new Runnable() {
			@Override
			public void run() {
				execute();
			}
		});

		return null;
	}
	protected void execute() {
		// TODO Auto-generated method stub
		shell = AIFDesktop.getActiveDesktop().getShell();

		ExistingSheetPagesDialog dialog = new ExistingSheetPagesDialog(shell, SWT.SHELL_TRIM, shList,"选择减页sheet");
		dialog.open();

		sheetlist = dialog.list;
		if (sheetlist == null || sheetlist.size()<1) {
			return;
		}
		
		Thread thread = new Thread() {
			public void run() {
				try {
					new RemovePagesStationInfoTableOp(session, book, topbomline,sheetlist,info);
				} catch (TCException  e) {
					// TODO Auto-generated catch block
					e.printStackTrace();
				}	
			}
		};
		thread.start();
	}

}
