package com.dfl.report.workschedule;

import java.io.IOException;
import java.io.InputStream;
import java.util.List;
import java.util.Map;

import javax.swing.JOptionPane;

import org.apache.poi.ss.formula.functions.T;
import org.eclipse.core.commands.AbstractHandler;
import org.eclipse.core.commands.ExecutionEvent;
import org.eclipse.core.commands.ExecutionException;
import org.eclipse.swt.widgets.Display;

import com.dfl.report.ExcelReader.baseinfoExcelReader;
import com.dfl.report.util.FileUtil;
import com.dfl.report.util.GenerateReportInfo;
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

public class BasicInformationBZHandler extends AbstractHandler {

	private AbstractAIFUIApplication app;
	private GenerateReportInfo info;
	private InputStream inputStream = null;
	private Map<String, List<String>> MaterialMap;

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
			MessageBox.post("��ѡ��һ�ĺ�װ�������ն���", "��ܰ��ʾ", MessageBox.INFORMATION);
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
		// ��ȡ���϶��ձ�
		MaterialMap = baseinfoExcelReader.getMaterialComparisonTable(app, "DFL_MaterialMapping");
		if (MaterialMap == null || MaterialMap.size() < 1) {
			System.out.println("δ�ҵ����϶��ձ�");
			MessageBox.post("δ���ö��ձ�DFL_MaterialMapping������ϵϵͳ����Ա��", "��ʾ��Ϣ", MessageBox.ERROR);
			return null;
		}

		System.out.println("���ڲ�������Ӧ����");

		/******************* �жϷŵ�ǰ�� *******************************/
		// �ļ�����
		String procName = "222.������Ϣ";
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

		if (info.getAction() == "create") { // �����
			MessageBox.post("�������������Ϣ-�����嵥��Ϣ��", "��Ϣ��ʾ", MessageBox.INFORMATION);
			return null;
		} else {
			TCComponentItemRevision docmentRev = info.getMeDocument();
			inputStream = baseinfoExcelReader.getFileinbyreadExcel(docmentRev, "IMAN_specification", procName);

			if (inputStream == null) {
				MessageBox.post("��ȷ��222.������Ϣ�ĵ��汾�����£�����222.������Ϣ���ݼ���", "��Ϣ��ʾ", MessageBox.INFORMATION);
				return null;
			}
		}
		/*************************************************/

		Thread thread = new Thread() {
			public void run() {
				boolean IsContinu = Util.isContinue("�Ḳ����һ������İ�����Ϣ����ȷ���Ƿ�����������");
				if (IsContinu) {
					try {
						new BasicInformationBZOp(app, info, inputStream,MaterialMap);
					} catch (TCException e) {
						// TODO Auto-generated catch block
						e.printStackTrace();
					}
				} else {
					try {
						inputStream.close();
					} catch (IOException e) {
						// TODO Auto-generated catch block
						e.printStackTrace();
					}
				}
			}
		};
		thread.start();

		return null;
	}
}
