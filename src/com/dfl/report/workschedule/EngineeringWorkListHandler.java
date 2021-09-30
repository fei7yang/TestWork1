package com.dfl.report.workschedule;

import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.List;

import org.eclipse.core.commands.AbstractHandler;
import org.eclipse.core.commands.ExecutionEvent;
import org.eclipse.core.commands.ExecutionException;
import org.eclipse.swt.SWT;
import org.eclipse.swt.widgets.Display;
import org.eclipse.swt.widgets.Shell;

import com.dfl.report.ExcelReader.CoverInfomation;
import com.dfl.report.ExcelReader.baseinfoExcelReader;
import com.dfl.report.util.FileUtil;
import com.dfl.report.util.GenerateReportInfo;
import com.dfl.report.util.ReportUtils;
import com.teamcenter.rac.aif.AIFDesktop;
import com.teamcenter.rac.aif.AbstractAIFUIApplication;
import com.teamcenter.rac.aif.kernel.InterfaceAIFComponent;
import com.teamcenter.rac.aifrcp.AIFUtility;
import com.teamcenter.rac.kernel.TCComponentBOMLine;
import com.teamcenter.rac.kernel.TCException;
import com.teamcenter.rac.util.MessageBox;

public class EngineeringWorkListHandler extends AbstractHandler {

	private AbstractAIFUIApplication app;
	private Shell shell;
	private String Edition;
	private String topfoldername;
	private GenerateReportInfo info;
	private InputStream inputStream = null;

	public EngineeringWorkListHandler() {
		// TODO Auto-generated constructor stub
	}

	@Override
	public Object execute(ExecutionEvent arg0) throws ExecutionException {
		// TODO Auto-generated method stub
		app = AIFUtility.getCurrentApplication();
		InterfaceAIFComponent[] ifc = app.getTargetComponents();
		if (ifc.length < 1) {
			MessageBox.post("��ǰδѡ�������������ѡ��", "��ܰ��ʾ", MessageBox.INFORMATION);
			return null;
		}
		if (ifc.length > 1) {
			MessageBox.post("��ѡ��һ��װ�������ն���", "��ܰ��ʾ", MessageBox.INFORMATION);
			return null;
		}
		if (ifc[0] instanceof TCComponentBOMLine) {

		} else {
			MessageBox.post("��ѡ��װ�������ն���", "��ʾ", MessageBox.INFORMATION);
			return null;
		}
		TCComponentBOMLine topbomline = (TCComponentBOMLine) ifc[0];
		try {
			System.out.println(topbomline.getItemRevision().getType());
			if (!topbomline.getItemRevision().isTypeOf("B8_BIWPlantBOPRevision")) {
				MessageBox.post("��ѡ��װ�������ն���", "��ܰ��ʾ", MessageBox.INFORMATION);
				return null;
			}
		} catch (TCException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}

		// �ļ�����
		String procName = "01.Ŀ¼";

		// ���ɱ������ǰ�Ķ���
		info = new GenerateReportInfo();
		info.setExist(false);
		info.setIsgoon(true);
		info.setAction(""); //$NON-NLS-1$
		info.setMeDocument(null);
		info.setDFL9_process_type("H"); //$NON-NLS-1$
		info.setDFL9_process_file_type("AB"); // $NON-NLS-1$
		info.setmeDocumentName(procName);

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
		// ��ѯĿ¼����ģ��
		inputStream = FileUtil.getTemplateFile("DFL_Template_EngineeringWorkListContents");

		if (inputStream == null) {
			MessageBox.post("����û���ҵ�������ҵ��Ŀ¼ģ�壬����ϵϵͳ����Ա���ģ��(����Ϊ��DFL_Template_EngineeringWorkListContents)", "��ܰ��ʾ", MessageBox.INFORMATION);
			return null;
		}

		List<CoverInfomation> list = getCoverinfomation(topbomline, "00.����");
		if (list != null && list.size() > 0) {
			CoverInfomation cif = list.get(0);
			Edition = cif.getEdition();
			topfoldername = cif.getFilecode();
		} else {
			try {
				inputStream.close();
			} catch (IOException e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			}
			MessageBox.post("�������ɹ�����ҵ��-������Ϣ��", "��ܰ��ʾ", MessageBox.INFORMATION);
			return null;
		}
		Thread thread = new Thread() {
			public void run() {
				try {
					new EngineeringWorkListOp(app, null, Edition, topfoldername,info,inputStream);
				} catch (TCException e) {
					// TODO Auto-generated catch block
					e.printStackTrace();
				}
			}
		};
		thread.start();


		return null;
	}

	protected void openDialog() {
		// TODO Auto-generated method stub
//		EditionDialog dialog = new EditionDialog(shell, SWT.SHELL_TRIM);
//		dialog.open();
//
//		Edition = dialog.Edition;
//		if (Edition.isEmpty()) {
//			return;
//		}
//		
//		Thread thread = new Thread() {
//			public void run() {
//				try {
//					new EngineeringWorkListOp(app, null, Edition);
//				} catch (TCException e) {
//					// TODO Auto-generated catch block
//					e.printStackTrace();
//				}	
//			}
//		};
//		thread.start();

	}

	/*
	 * ��ȡ������Ϣ��Ϣ
	 */
	private List<CoverInfomation> getCoverinfomation(TCComponentBOMLine topbl, String procName) {
		List<CoverInfomation> coverinfolist = new ArrayList<CoverInfomation>();
		InputStream filein = null;
		try {
			filein = baseinfoExcelReader.getFileinbyreadExcel2(topbl.getItemRevision(), "IMAN_reference", procName);
		} catch (TCException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		coverinfolist = baseinfoExcelReader.readCoverExcel(filein, "xlsx");

		return coverinfolist;
	}
}
