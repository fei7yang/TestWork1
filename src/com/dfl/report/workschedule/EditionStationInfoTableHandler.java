package com.dfl.report.workschedule;

import java.io.InputStream;
import java.util.ArrayList;
import java.util.List;

import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.eclipse.core.commands.AbstractHandler;
import org.eclipse.core.commands.ExecutionEvent;
import org.eclipse.core.commands.ExecutionException;
import org.eclipse.swt.widgets.Shell;

import com.dfl.report.ExcelReader.CoverInfomation;
import com.dfl.report.ExcelReader.baseinfoExcelReader;
import com.dfl.report.util.GenerateReportInfo;
import com.dfl.report.util.NewOutputDataToExcel;
import com.dfl.report.util.ReportUtils;
import com.dfl.report.util.Util;
import com.teamcenter.rac.aif.AbstractAIFUIApplication;
import com.teamcenter.rac.aif.kernel.InterfaceAIFComponent;
import com.teamcenter.rac.aifrcp.AIFUtility;
import com.teamcenter.rac.kernel.TCComponentBOMLine;
import com.teamcenter.rac.kernel.TCComponentItemRevision;
import com.teamcenter.rac.kernel.TCException;
import com.teamcenter.rac.kernel.TCSession;
import com.teamcenter.rac.util.MessageBox;

public class EditionStationInfoTableHandler extends AbstractHandler {

	private AbstractAIFUIApplication app;
	private TCSession session;
	private XSSFWorkbook book;
	private TCComponentBOMLine topbomline;
	private GenerateReportInfo info;
	public String Edition = "";
	@Override
	public Object execute(ExecutionEvent arg0) throws ExecutionException {
		// TODO Auto-generated method stub
		app = AIFUtility.getCurrentApplication();
		session = (TCSession) app.getSession();
		InterfaceAIFComponent[] ifc = app.getTargetComponents();
		if (ifc.length < 1) {
			MessageBox.post("��ǰδѡ�������������ѡ��", "��ܰ��ʾ", MessageBox.INFORMATION);
			return null;
		}
		if (ifc.length > 1) {
			MessageBox.post("��ѡ��һ��װ��λ���ն���", "��ܰ��ʾ", MessageBox.INFORMATION);
			return null;
		}
		if (ifc[0] instanceof TCComponentBOMLine) {

		} else {
			MessageBox.post("��ѡ��װ��λ���ն���", "��ʾ", MessageBox.INFORMATION);
			return null;
		}
		topbomline = (TCComponentBOMLine) ifc[0];

		try {
			System.out.println(topbomline.getItemRevision().getType());
			if (!topbomline.getItemRevision().isTypeOf("B8_BIWMEProcStatRevision")) {
				MessageBox.post("��ѡ��װ��λ���ն���", "��ܰ��ʾ", MessageBox.INFORMATION);
				return null;
			}
		} catch (TCException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
//		List<CoverInfomation> list;
//		try {
//			list = getCoverinfomation(topbomline.window().getTopBOMLine(), "00.����");
//			if (list != null && list.size() > 0) {
//				CoverInfomation cif = list.get(0);
//				Edition = cif.getEdition();
//			} else {
//				System.out.println("�������ɹ�����ҵ��-���棡");
//				MessageBox.post("�������ɹ�����ҵ��-���棡", "��ʾ��Ϣ", MessageBox.ERROR);
//				return null;
//			}
//		} catch (TCException e1) {
//			// TODO Auto-generated catch block
//			e1.printStackTrace();
//		}
		
		// ��ȡ�����ɵı���
		info = new GenerateReportInfo();
		info.setExist(false);
		info.setIsgoon(true);
		info.setAction(""); //$NON-NLS-1$
		info.setMeDocument(null);
		info.setDFL9_process_type("H"); //$NON-NLS-1$
		info.setDFL9_process_file_type("AB"); // $NON-NLS-1$
		info.setmeDocumentName("");
		info.setFlag(true);
		try {
			info.setProject_ids(topbomline.window().getTopBOMLine().getItemRevision());
		} catch (TCException e1) {
			// TODO Auto-generated catch block
			e1.printStackTrace();
		}

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
			MessageBox.post("��ȷ���Ѿ����ɹ�����ҵ��", "��ܰ��ʾ", MessageBox.INFORMATION);
			return null;
		}
		String procName = Util.getProperty(docmentRev, "object_name");
		InputStream inputStream = baseinfoExcelReader.getFileinbyreadExcel(docmentRev, "IMAN_specification", procName);

		if (inputStream == null) {
			MessageBox.post("��ȷ��" + procName + "�汾�����£�����" + procName + "���ݼ���", "��ܰ��ʾ", MessageBox.INFORMATION);
			return null;
		}
		book = NewOutputDataToExcel.creatXSSFWorkbook(inputStream);
		
		Thread thread = new Thread() {
			public void run() {
				try {
					new EditionStationInfoTableOp(session, book, topbomline,info,Edition);
				} catch (TCException  e) {
					// TODO Auto-generated catch block
					e.printStackTrace();
				}	
			}
		};
		thread.start();
		
		return null;
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
