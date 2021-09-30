package com.dfl.report.workschedule;

import java.io.InputStream;
import java.rmi.AccessException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;

import org.eclipse.core.commands.AbstractHandler;
import org.eclipse.core.commands.ExecutionEvent;
import org.eclipse.core.commands.ExecutionException;
import org.eclipse.swt.widgets.Shell;

import com.dfl.report.ExcelReader.CoverInfomation;
import com.dfl.report.ExcelReader.CurrentandVoltage;
import com.dfl.report.ExcelReader.WeldPointBoardInformation;
import com.dfl.report.ExcelReader.baseinfoExcelReader;
import com.dfl.report.util.FileUtil;
import com.dfl.report.util.GenerateReportInfo;
import com.dfl.report.util.ReportUtils;
import com.teamcenter.rac.aif.AbstractAIFUIApplication;
import com.teamcenter.rac.aif.kernel.InterfaceAIFComponent;
import com.teamcenter.rac.aifrcp.AIFUtility;
import com.teamcenter.rac.kernel.TCComponentBOMLine;
import com.teamcenter.rac.kernel.TCComponentItemRevision;
import com.teamcenter.rac.kernel.TCComponentItemType;
import com.teamcenter.rac.kernel.TCException;
import com.teamcenter.rac.kernel.TCPreferenceService;
import com.teamcenter.rac.kernel.TCSession;
import com.teamcenter.rac.util.MessageBox;

public class AlterStationInfoTableHandler extends AbstractHandler {

	private AbstractAIFUIApplication app;
	private String Edition;
	private String topfoldername;
	private TCSession session;
	private List<CurrentandVoltage> cv;
	private List<WeldPointBoardInformation> baseinfolist;
	private Map<String, List<String>> MaterialMap;

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
		TCComponentBOMLine topbomline = (TCComponentBOMLine) ifc[0];

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
		try {
			System.out.println(topbomline.getItemRevision().getType());
			if (!topbomline.getItemRevision().isTypeOf("B8_BIWMEProcStatRevision")) {
				MessageBox.post("��ѡ��װ��λ���ն���", "��ʾ��Ϣ", MessageBox.INFORMATION);
				return null;
			}
		} catch (TCException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		// ��ȡ��ѡ����Note����
		TCPreferenceService ts = session.getPreferenceService();
		if (!ts.isDefinitionExistForPreference("B8_Calculation_Parameter_Name")) {
			MessageBox.post("������ѡ��B8_Calculation_Parameter_Nameδ����,����ϵϵͳ����Ա��", "��ʾ��Ϣ", MessageBox.ERROR);
			return null;
		}
		// ��ȡ���϶��ձ�
		MaterialMap = baseinfoExcelReader.getMaterialComparisonTable(app, "DFL_MaterialMapping");
		if (MaterialMap == null || MaterialMap.size() < 1) {
			System.out.println("δ�ҵ����϶��ձ�");
			MessageBox.post("δ���ö��ձ�DFL_MaterialMapping������ϵϵͳ����Ա��", "��ʾ��Ϣ", MessageBox.ERROR);
			return null;
		}
		TCComponentItemRevision toprev = null;
		try {
			toprev = topbomline.window().getTopBOMLine().getItemRevision();
		} catch (TCException e1) {
			// TODO Auto-generated catch block
			e1.printStackTrace();
		}
		// ��ȡ�����ɵı���
		GenerateReportInfo info = new GenerateReportInfo();
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
		if (docmentRev == null) {
			MessageBox.post("��ȷ���Ѿ����ɹ�����ҵ��", "��ܰ��ʾ", MessageBox.INFORMATION);
			return null;
		}
		// ��ȡ�������
		Object[] obj = baseinfoExcelReader.getCalculationParameter(app, "B8_Calculation_Parameter_Name");
		cv = new ArrayList<CurrentandVoltage>();
		if (obj != null) {
			if (obj[1] != null) {
				cv = (List<CurrentandVoltage>) obj[1];
			} else {
				System.out.println("δ�ҵ����Ӳ����������");
				MessageBox.post("����δ�ҵ����Ӳ����������", "��ʾ��Ϣ", MessageBox.ERROR);
				return null;
			}
		} else {
			System.out.println("δ�ҵ����Ӳ����������");
			MessageBox.post("����δ�ҵ����Ӳ����������", "��ʾ��Ϣ", MessageBox.ERROR);
			return null;
		}
		Map<String, String> map = getSizeRule();
		if (map == null || map.size() < 1) {
			System.out.println("��ѡ��DFL9_get_parts_source δ���ã�����ϵϵͳ����Ա��");
			MessageBox.post("������ѡ��DFL9_get_parts_source δ���ã�����ϵϵͳ����Ա��", "��ʾ��Ϣ", MessageBox.ERROR);
			return null;
		}
		String baseName = "222.������Ϣ";
		try {
			baseinfolist = getBaseinfomation(topbomline.window().getTopBOMLine(), baseName);
			if (baseinfolist == null || baseinfolist.size() < 1) {
				System.out.println("�������ɹ�����ҵ��-������Ϣ��");
				MessageBox.post("�������ɹ�����ҵ��-������Ϣ��", "��ʾ��Ϣ", MessageBox.ERROR);
				return null;
			}
		} catch (TCException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		List<CoverInfomation> list;
		try {
			list = getCoverinfomation(topbomline.window().getTopBOMLine(), "00.����");
			if (list != null && list.size() > 0) {
				CoverInfomation cif = list.get(0);
				Edition = cif.getEdition();
				topfoldername = cif.getFilecode();
			} else {
				System.out.println("�������ɹ�����ҵ��-���棡");
				MessageBox.post("�������ɹ�����ҵ��-���棡", "��ʾ��Ϣ", MessageBox.ERROR);
				return null;
			}
		} catch (TCException e1) {
			// TODO Auto-generated catch block
			e1.printStackTrace();
		}

		Thread thread = new Thread() {
			public void run() {
				try {
					new StationInformationTableOp(app, Edition, topfoldername, cv, baseinfolist,MaterialMap);
				} catch (TCException | AccessException e) {
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

	// ��ѯ��Ʒ������ѡ���ȡ��Ʒ������Ϣ
	private Map<String, String> getSizeRule() {
		Map<String, String> rule = new HashMap<String, String>();
		try {
			AbstractAIFUIApplication app = AIFUtility.getCurrentApplication();
			TCSession session = (TCSession) app.getSession();
			TCPreferenceService preferenceService = session.getPreferenceService();

			String str = preferenceService.getPreferenceDescription("DFL9_get_parts_source");
			if (str != null) {
				String[] values = preferenceService.getStringValues("DFL9_get_parts_source");
				for (int i = 0; i < values.length; i++) {
					String value = values[i];
					if (value != null) {
						String[] val = value.split("=");
						if (val != null && val.length > 1) {
							rule.put(val[0], val[1]);
						}
					}
				}
			}
			return rule;
		} catch (Exception e) {
			e.printStackTrace();
		}
		return rule;
	}

	/*
	 * ��ȡ������Ϣ����Ϣ
	 */
	private List<WeldPointBoardInformation> getBaseinfomation(TCComponentBOMLine topbl, String procName) {
		List<WeldPointBoardInformation> baseinfolist = new ArrayList<WeldPointBoardInformation>();
		InputStream filein = null;
		try {
			filein = baseinfoExcelReader.getFileinbyreadExcel2(topbl.getItemRevision(), "IMAN_reference", procName);
		} catch (TCException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		baseinfolist = baseinfoExcelReader.readHDExcel(filein, "xlsx");

		return baseinfolist;
	}
}
