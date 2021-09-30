package com.dfl.report.workschedule;

import java.io.InputStream;
import java.math.BigDecimal;
import java.text.DecimalFormat;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.logging.Logger;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.eclipse.core.commands.AbstractHandler;
import org.eclipse.core.commands.ExecutionEvent;
import org.eclipse.core.commands.ExecutionException;
import org.eclipse.swt.SWT;
import org.eclipse.swt.widgets.Display;
import org.eclipse.swt.widgets.Shell;

import com.dfl.report.ExcelReader.CoverInfomation;
import com.dfl.report.ExcelReader.CurrentandVoltage;
import com.dfl.report.ExcelReader.RecommendedPressure;
import com.dfl.report.ExcelReader.SFSequenceWeldingConditionList;
import com.dfl.report.ExcelReader.SequenceComparisonTable;
import com.dfl.report.ExcelReader.SequenceWeldingConditionList;
import com.dfl.report.ExcelReader.WeldPointBoardInformation;
import com.dfl.report.ExcelReader.baseinfoExcelReader;
import com.dfl.report.common.TCComponentUtils;
import com.dfl.report.dealparameter.DealParameterHandler;
import com.dfl.report.util.FileUtil;
import com.dfl.report.util.GenerateReportInfo;
import com.dfl.report.util.NewOutputDataToExcel;
import com.dfl.report.util.ReportUtils;
import com.dfl.report.util.ReportViwePanel;
import com.dfl.report.util.Util;
import com.teamcenter.rac.aif.AIFDesktop;
import com.teamcenter.rac.aif.AbstractAIFUIApplication;
import com.teamcenter.rac.aif.kernel.InterfaceAIFComponent;
import com.teamcenter.rac.aifrcp.AIFUtility;
import com.teamcenter.rac.kernel.TCComponent;
import com.teamcenter.rac.kernel.TCComponentBOMLine;
import com.teamcenter.rac.kernel.TCComponentDataset;
import com.teamcenter.rac.kernel.TCComponentItem;
import com.teamcenter.rac.kernel.TCComponentItemRevision;
import com.teamcenter.rac.kernel.TCException;
import com.teamcenter.rac.kernel.TCPreferenceService;
import com.teamcenter.rac.kernel.TCSession;
import com.teamcenter.rac.kernel.TCUserService;
import com.teamcenter.rac.util.MessageBox;

/* *****************************************
 * ���¹�����ҵ��ֻ�����ڼ����û�ά����������Ĳ���ֵ
 * @hgq
 * 20191026
 */
public class UpdateEngineeringWorkListHandler extends AbstractHandler {

	private AbstractAIFUIApplication app;
	private static Logger logger = Logger.getLogger(UpdateEngineeringWorkListHandler.class.getName()); // ��־��ӡ��
	private List<WeldPointBoardInformation> baseinfolist = new ArrayList<WeldPointBoardInformation>();// ������Ϣ�������
	private static List<SequenceWeldingConditionList> swc = new ArrayList<SequenceWeldingConditionList>();// 24���к��������趨��
																											// ���к�
	private static List<CurrentandVoltage> cv = new ArrayList<CurrentandVoltage>();// 24���к��������趨�� ������ѹ
	private static List<SFSequenceWeldingConditionList> SFswc = new ArrayList<SFSequenceWeldingConditionList>();// 255���к��������趨��
	private static List<RecommendedPressure> rp = new ArrayList<RecommendedPressure>();// �Ƽ���ѹ��
	private static List<SequenceComparisonTable> sct = new ArrayList<SequenceComparisonTable>();// ���ж��ձ�
	private TCSession session;
	private Map<String, List<String>> MaterialMap;
	private Shell shell;
	private TCComponentBOMLine topbomline;
	private String result;

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

		// ��ȡ��ѡ����Note����
		TCPreferenceService ts = session.getPreferenceService();
		if (!ts.isDefinitionExistForPreference("B8_Calculation_Parameter_Name")) {
			MessageBox.post("������ѡ��B8_Calculation_Parameter_Nameδ���壬����ϵϵͳ����Ա��", "��ʾ��Ϣ", MessageBox.ERROR);
			return null;
		}
		// ��ȡ���϶��ձ�
		MaterialMap = baseinfoExcelReader.getMaterialComparisonTable(app, "DFL_MaterialMapping");
		if (MaterialMap == null || MaterialMap.size() < 1) {
			System.out.println("δ�ҵ����϶��ձ�");
			MessageBox.post("δ���ö��ձ�DFL_MaterialMapping������ϵϵͳ����Ա��", "��ʾ��Ϣ", MessageBox.ERROR);
			return null;
		}

		shell = AIFDesktop.getActiveDesktop().getShell();
		Display.getDefault().asyncExec(new Runnable() {
			@Override
			public void run() {
				openDialog();
			}
		});

		return null;
	}

	protected void openDialog() {
		// TODO Auto-generated method stub
		DetermineDialog dialog = new DetermineDialog(shell, SWT.SHELL_TRIM);
		dialog.open();

		result = dialog.getMessage();

		Thread thread = new Thread() {
			public void run() {
				if (!result.isEmpty()) {
					try {
						UpdateEngineeringWorkList(topbomline, result);
					} catch (TCException e) {
						// TODO Auto-generated catch block
						e.printStackTrace();
					}
				}
			}
		};
		thread.start();

	}

	private void UpdateEngineeringWorkList(TCComponentBOMLine topbomline, String result) throws TCException {
		// TODO Auto-generated method stub

		String procName = "";
		// ���ɱ������ǰ�Ķ���
		GenerateReportInfo info = new GenerateReportInfo();
		info.setExist(false);
		info.setIsgoon(true);
		info.setAction(""); //$NON-NLS-1$
		info.setMeDocument(null);
		info.setDFL9_process_type("H"); //$NON-NLS-1$
		info.setDFL9_process_file_type("AB"); // $NON-NLS-1$
		info.setmeDocumentName(procName);
		info.setFlag(true);
		info.setProject_ids(topbomline.window().getTopBOMLine().getItemRevision());

		try {
			info = ReportUtils.beforeGenerateReportAction(topbomline.getItemRevision(), info);
		} catch (TCException e) {
			e.printStackTrace();
			// EclipseUtils.info("Error : " + e.getMessage()); //$NON-NLS-1$
			return;
		}
		System.out.println("The action is completed before the report operation is generated.");

		if (!info.isIsgoon()) {
			return;
		}
		if (!info.isExist()) {
			MessageBox.post("��ȷ���Ѿ����ɹ�����ҵ��", "��ܰ��ʾ", MessageBox.INFORMATION);
			return;
		}
		InputStream inputStream = null;
		TCComponentItemRevision docmentRev = info.getMeDocument();
		procName = Util.getProperty(docmentRev, "object_name");
		inputStream = baseinfoExcelReader.getFileinbyreadExcel(docmentRev, "IMAN_specification", procName);

		if (inputStream == null) {
			MessageBox.post("��ȷ��" + procName + "�汾�����£�����" + procName + "���ݼ���", "��ܰ��ʾ", MessageBox.INFORMATION);
			return;
		}
		// ��ȡ������Ϣ
		String baseName = "222.������Ϣ";
		TCComponentBOMLine topbl = topbomline.window().getTopBOMLine();
		baseinfolist = getBaseinfomation(topbl, baseName);
		if (baseinfolist == null || baseinfolist.size() < 1) {
			System.out.println("�������ɹ�����ҵ��-������Ϣ��");
			MessageBox.post("�������ɹ�����ҵ��-������Ϣ��", "��ʾ��Ϣ", MessageBox.ERROR);
			return;
		}
		// ��ȡ�������
		Object[] obj = baseinfoExcelReader.getCalculationParameter(app, "B8_Calculation_Parameter_Name");
		if (obj != null) {
			if (obj[0] != null) {
				swc = (List<SequenceWeldingConditionList>) obj[0];
			} else {
				System.out.println("δ��ȡ��24���к��������趨�� ���к���Ϣ��");
				MessageBox.post("δ�ҵ����Ӳ����������", "��ʾ��Ϣ", MessageBox.ERROR);
				return;
			}
			if (obj[1] != null) {
				cv = (List<CurrentandVoltage>) obj[1];
			} else {
				System.out.println("δ��ȡ��24���к��������趨�� ������ѹ��Ϣ��");
				MessageBox.post("δ�ҵ����Ӳ����������", "��ʾ��Ϣ", MessageBox.ERROR);
				return;
			}
			if (obj[2] != null) {
				SFswc = (List<SFSequenceWeldingConditionList>) obj[2];
			} else {
				System.out.println("δ��ȡ��255���к��������趨����Ϣ��");
				MessageBox.post("δ�ҵ����Ӳ����������", "��ʾ��Ϣ", MessageBox.ERROR);
				return;
			}
			if (obj[3] != null) {
				rp = (List<RecommendedPressure>) obj[3];
			} else {
				System.out.println("δ��ȡ���Ƽ���ѹ����Ϣ��");
				MessageBox.post("δ�ҵ����Ӳ����������", "��ʾ��Ϣ", MessageBox.ERROR);
				return;
			}
			if (obj[4] != null) {
				sct = (List<SequenceComparisonTable>) obj[4];
			} else {
				System.out.println("δ��ȡ�����ж��ձ���Ϣ��");
				MessageBox.post("δ�ҵ����Ӳ����������", "��ʾ��Ϣ", MessageBox.ERROR);
				return;
			}
		} else {
			MessageBox.post("δ�ҵ����Ӳ����������", "��ʾ��Ϣ", MessageBox.ERROR);
			return;
		}

		// ��ʾ�����������
		// ������ʾ���Ȳ����ִ�в���
		ReportViwePanel viewPanel = new ReportViwePanel("�������");
		viewPanel.setVisible(true);
		viewPanel.addInfomation("���ڸ���...\n", 5, 100);

		XSSFWorkbook book = NewOutputDataToExcel.creatXSSFWorkbook(inputStream);

		viewPanel.addInfomation("", 20, 100);

		// ������·
		{
			Util.callByPass(session, true);
		}
		// ����PSWsheet ҳ
		DealPSWSheet(book, viewPanel, result);

		viewPanel.addInfomation("���ڸ���...\n", 40, 100);

		// ����RSW����
		DealRSWQDSheet(book, viewPanel);

		// ����RSW�ŷ�
		DealRSWSFSheet(book, viewPanel);

		viewPanel.addInfomation("", 60, 100);

		TCComponentItemRevision dfl9MEDocumentRev = info.getMeDocument();
		TCComponentDataset tagedataset = null;
		TCComponent[] children = TCComponentUtils.getCompsByRelation(dfl9MEDocumentRev, "IMAN_specification");
		for (TCComponent child : children) {
			if (child instanceof TCComponentDataset) {
				TCComponentDataset dataset = (TCComponentDataset) child;
				tagedataset = dataset;
				break;
			}
		}
		String fileName = Util.formatString(Util.getProperty(tagedataset, "object_name"));
		// ����ļ�
		NewOutputDataToExcel.exportFile(book, fileName);

		viewPanel.addInfomation("", 80, 100);

		String fullFileName = FileUtil.getReportFileName(fileName);
		System.out.println(fullFileName);
		TCComponentDataset ds = Util.createDataset(session, fileName, fullFileName, "MSExcelX", "excel");
		List<TCComponentDataset> datasetList = new ArrayList<TCComponentDataset>();
		List<TCComponentItemRevision> revlist = new ArrayList<TCComponentItemRevision>();
		if (ds != null) {
			datasetList.add(ds);
		}
		revlist.add(topbomline.getItemRevision());
		try {
			TCComponentItem docunment = ReportUtils.afterGenerateReportAction(datasetList, revlist, info, procName, "",
					session);
			// saveFileToFolder(docunment, topfoldername, childrenFoldername);

		} catch (TCException e) {
			e.printStackTrace();
			// EclipseUtils.info(Messages.FixtureOperation_15 + e.getMessage());
			return;
		}
		// �ر���·
		{
			Util.callByPass(session, false);
		}
		viewPanel.addInfomation("�����������ݸ�����ɡ�", 100, 100);
	}

	private void DealRSWSFSheet(XSSFWorkbook book, ReportViwePanel viewPanel) {
		// TODO Auto-generated method stub
		int sheetnum = 0;
		sheetnum = book.getNumberOfSheets();
		ArrayList sheetAtIndexs = new ArrayList(); // RSW��������λ��
		for (int i = 0; i < sheetnum; i++) {
			String sheetname = book.getSheetName(i);
			if (sheetname.contains("RSW�ŷ�")) {
				sheetAtIndexs.add(i);
			}
		}
		if (sheetAtIndexs == null && sheetAtIndexs.size() < 1) {
			return;
		}
		// ��ȡsheet�ڵ�����
		ArrayList datalist = getSheetData(book, sheetAtIndexs, true);

		ArrayList hdlist = new ArrayList();

		// ���ݺ����ڻ�����Ϣ���л�ȡ�����Ϣ
		List hdinfo = new ArrayList();// ����������Ϣ

		hdinfo = getBoardInformation(baseinfolist, datalist);

		// ѭ��������Ϣ�����㲢��ȡ��������ֵ
		Map<String, String[]> paramap = getDealParameter(hdinfo, true, viewPanel);

		// �������������λ�ã�����ȡ������Ϣд�뵽������
		writeRSWSFInfomation(book, hdinfo, paramap);
	}

	private void writeRSWSFInfomation(XSSFWorkbook book, List hdinfo, Map<String, String[]> paramap) {
		// TODO Auto-generated method stub
		if (hdinfo != null && hdinfo.size() > 0) {
			// ����������ɫ
			Font font = book.createFont();
			font.setColor((short) 12);// ��ɫ����
			font.setFontHeightInPoints((short) 9);
			XSSFCellStyle style = book.createCellStyle();
			style.setBorderBottom(XSSFCellStyle.BORDER_THIN); // �±߿�
			style.setBorderLeft(XSSFCellStyle.BORDER_THIN);// ��߿�
			style.setBorderTop(XSSFCellStyle.BORDER_THIN);// �ϱ߿�
			style.setBorderRight(XSSFCellStyle.BORDER_THIN);// �ұ߿�
			style.setVerticalAlignment(XSSFCellStyle.VERTICAL_CENTER);
			style.setAlignment(XSSFCellStyle.ALIGN_CENTER);
			style.setFont(font);

			Font font2 = book.createFont();
			font2.setColor((short) 12);// ��ɫ����
			font2.setFontHeightInPoints((short) 18);
			font2.setBoldweight(Font.BOLDWEIGHT_BOLD); // ����Ӵ�
			XSSFCellStyle style2 = book.createCellStyle();
			style2.setBorderBottom(CellStyle.BORDER_MEDIUM); // ���߱߿�
			style2.setBorderLeft(CellStyle.BORDER_MEDIUM); // ���߱߿�
			style2.setBorderRight(CellStyle.BORDER_MEDIUM); // ���߱߿�
			style2.setBorderTop(CellStyle.BORDER_MEDIUM); // ���߱߿�
			style2.setVerticalAlignment(XSSFCellStyle.VERTICAL_CENTER);
			style2.setAlignment(XSSFCellStyle.ALIGN_CENTER);
			style2.setFont(font2);

			// ��ɫ����ɫ
			XSSFCellStyle style3 = book.createCellStyle();
			style3.setFillForegroundColor(IndexedColors.PINK.getIndex());
			style3.setFillPattern(CellStyle.SOLID_FOREGROUND);
			style3.setBorderBottom(XSSFCellStyle.BORDER_THIN); // �±߿�
			style3.setBorderLeft(XSSFCellStyle.BORDER_THIN);// ��߿�
			style3.setBorderTop(XSSFCellStyle.BORDER_THIN);// �ϱ߿�
			style3.setBorderRight(XSSFCellStyle.BORDER_THIN);// �ұ߿�
			style3.setVerticalAlignment(XSSFCellStyle.VERTICAL_CENTER);
			style3.setAlignment(XSSFCellStyle.ALIGN_CENTER);
			style3.setFont(font);
			// ��ɫ����ɫ
			Font font3 = book.createFont();
			font3.setColor((short) 1);// ��ɫ����
			font3.setFontHeightInPoints((short) 10);
			XSSFCellStyle style4 = book.createCellStyle();
			style4.setFillForegroundColor(IndexedColors.VIOLET.getIndex());
			style4.setFillPattern(CellStyle.SOLID_FOREGROUND);
			style4.setBorderBottom(XSSFCellStyle.BORDER_THIN); // �±߿�
			style4.setBorderLeft(XSSFCellStyle.BORDER_THIN);// ��߿�
			style4.setBorderTop(XSSFCellStyle.BORDER_THIN);// �ϱ߿�
			style4.setBorderRight(XSSFCellStyle.BORDER_THIN);// �ұ߿�
			style4.setVerticalAlignment(XSSFCellStyle.VERTICAL_CENTER);
			style4.setAlignment(XSSFCellStyle.ALIGN_CENTER);
			style4.setFont(font3);
			// ��ɫ����ɫ
			Font font4 = book.createFont();
			font4.setColor((short) 1);// ��ɫ����
			font4.setFontHeightInPoints((short) 10);
			XSSFCellStyle style5 = book.createCellStyle();
			style5.setFillForegroundColor(IndexedColors.SKY_BLUE.getIndex());
			style5.setFillPattern(CellStyle.SOLID_FOREGROUND);
			style5.setBorderBottom(XSSFCellStyle.BORDER_THIN); // �±߿�
			style5.setBorderLeft(XSSFCellStyle.BORDER_THIN);// ��߿�
			style5.setBorderTop(XSSFCellStyle.BORDER_THIN);// �ϱ߿�
			style5.setBorderRight(XSSFCellStyle.BORDER_THIN);// �ұ߿�
			style5.setVerticalAlignment(XSSFCellStyle.VERTICAL_CENTER);
			style5.setAlignment(XSSFCellStyle.ALIGN_CENTER);
			style5.setFont(font4);

			XSSFCellStyle style6 = book.createCellStyle();
			style6.setBorderBottom(XSSFCellStyle.BORDER_THIN); // �±߿�
			// style6.setBorderLeft(XSSFCellStyle.BORDER_THIN);// ��߿�
			style6.setBorderTop(XSSFCellStyle.BORDER_THIN);// �ϱ߿�
			style6.setBorderRight(XSSFCellStyle.BORDER_THIN);// �ұ߿�
			style6.setVerticalAlignment(XSSFCellStyle.VERTICAL_CENTER);
			style6.setAlignment(XSSFCellStyle.ALIGN_CENTER);
			style6.setFont(font);

			XSSFCellStyle style8 = book.createCellStyle();
			style8.setBorderBottom(XSSFCellStyle.BORDER_THIN); // �±߿�
			style8.setBorderLeft(XSSFCellStyle.BORDER_THIN);// ��߿�
			style8.setBorderTop(XSSFCellStyle.BORDER_THIN);// �ϱ߿�
			// style8.setBorderRight(XSSFCellStyle.BORDER_THIN);// �ұ߿�
			style8.setVerticalAlignment(XSSFCellStyle.VERTICAL_CENTER);
			style8.setAlignment(XSSFCellStyle.ALIGN_CENTER);
			style8.setFont(font);

			// ��ɫ����ɫ
			Font font5 = book.createFont();
			font4.setFontHeightInPoints((short) 10);
			XSSFCellStyle style7 = book.createCellStyle();
			style7.setFillForegroundColor(IndexedColors.WHITE.getIndex());
			style7.setFillPattern(CellStyle.SOLID_FOREGROUND);
			style7.setBorderBottom(XSSFCellStyle.BORDER_THIN); // �±߿�
			style7.setBorderLeft(XSSFCellStyle.BORDER_THIN);// ��߿�
			style7.setBorderTop(XSSFCellStyle.BORDER_THIN);// �ϱ߿�
			style7.setBorderRight(XSSFCellStyle.BORDER_THIN);// �ұ߿�
			style7.setVerticalAlignment(XSSFCellStyle.VERTICAL_CENTER);
			style7.setAlignment(XSSFCellStyle.ALIGN_CENTER);
			style7.setFont(font5);

			// ����������ɫ
			Font font6 = book.createFont();
			font6.setColor((short) 2);// ��ɫ����
			font6.setFontHeightInPoints((short) 10);
			XSSFCellStyle style66 = book.createCellStyle();
			style66.setBorderBottom(XSSFCellStyle.BORDER_THIN); // �±߿�
			// style6.setBorderLeft(XSSFCellStyle.BORDER_THIN);// ��߿�
			style66.setBorderTop(XSSFCellStyle.BORDER_THIN);// �ϱ߿�
			style66.setBorderRight(XSSFCellStyle.BORDER_THIN);// �ұ߿�
			style66.setVerticalAlignment(XSSFCellStyle.VERTICAL_CENTER);
			style66.setAlignment(XSSFCellStyle.ALIGN_CENTER);
			style66.setFont(font6);

			// ��ɫ����ɫ
			Font fontpink = book.createFont();
			fontpink.setColor((short) 12);// ��ɫ����
			fontpink.setFontName("MS PGothic");
			fontpink.setFontHeightInPoints((short) 9);

			XSSFCellStyle stylepink = book.createCellStyle();
			stylepink.setFillForegroundColor(IndexedColors.ROSE.getIndex());
			stylepink.setFillPattern(CellStyle.SOLID_FOREGROUND);
			stylepink.setBorderBottom(XSSFCellStyle.BORDER_THIN); // �±߿�
			stylepink.setBorderLeft(XSSFCellStyle.BORDER_THIN);// ��߿�
			stylepink.setBorderTop(XSSFCellStyle.BORDER_THIN);// �ϱ߿�
			stylepink.setBorderRight(XSSFCellStyle.BORDER_THIN);// �ұ߿�
			stylepink.setVerticalAlignment(XSSFCellStyle.VERTICAL_CENTER);
			stylepink.setAlignment(XSSFCellStyle.ALIGN_CENTER);
			stylepink.setFont(fontpink);

			for (int i = 0; i < hdinfo.size(); i++) {
				String[] vals = (String[]) hdinfo.get(i);
				int sheetindex = Integer.parseInt(vals[2]); // sheet����λ��
				int rowindex = Integer.parseInt(vals[1]); // ��������
				XSSFSheet sheet = book.getSheetAt(sheetindex);

				String weldno = vals[3]; // ������
				String importance = vals[4]; // ��Ҫ��
				String boardnumber1 = vals[5]; // ���1���
				String boardname1 = vals[6]; // ���1����
				String partmaterial1 = vals[7]; // ���1����
				String partthickness1 = vals[8]; // ���1���
				String boardnumber2 = vals[9]; // ���2���
				String boardname2 = vals[10]; // ���2����
				String partmaterial2 = vals[11]; // ���2����
				String partthickness2 = vals[12]; // ���2���
				String boardnumber3 = vals[13]; // ���3���
				String boardname3 = vals[14]; // ���3����
				String partmaterial3 = vals[15]; // ���3����
				String partthickness3 = vals[16]; // ���3���
				String layersnum = vals[17]; // �����
				String gagi = vals[18]; // GA /GI
				String sheetstrength440 = vals[19]; // ����ǿ��(Mpa)440
				String sheetstrength590 = vals[20]; // ����ǿ��(Mpa)590
				String sheetstrength = vals[21]; // ����ǿ��(Mpa)>590
				String basethickness = vals[22]; // ��׼���
				String sheetstrength12 = vals[23]; // ����ǿ��(Mpa)1.2G
				String CurrentSerie = ""; // ���� ���� (�ղ�)
				String RecomWeldForce = "";// �Ƽ� ��ѹ��(N)
				String CurrentSeriedfi = ""; // ���� ���� (��Ӧ)

				// ���ݲ��ʶ��ձ��жϺ����Ƿ������㺸�Ӳ���
				boolean partmaterialFlag1 = true;
				boolean partmaterialFlag2 = true;
				boolean partmaterialFlag3 = true;
				String gagi1 = vals[27];
				String gagi2 = vals[28];
				String gagi3 = vals[29];

				// ���ݲ��ʶ��ձ��ȡGA/GI����
				if (MaterialMap != null) {
					for (Map.Entry<String, List<String>> entry : MaterialMap.entrySet()) {
						String MaterialNo = entry.getKey();
						List<String> infolist = entry.getValue();
						if (!"GA".equals(gagi1) && !"GI".equals(gagi1)) {
							if (Util.getIsEqueal(partmaterial1, MaterialNo)) {
								if ("��".equalsIgnoreCase(infolist.get(1))) {
									partmaterialFlag1 = false;
								}
							}
						}
						if (!"GA".equals(gagi2) && !"GI".equals(gagi2)) {
							if (Util.getIsEqueal(partmaterial2, MaterialNo)) {
								if ("��".equalsIgnoreCase(infolist.get(1))) {
									partmaterialFlag2 = false;
								}
							}
						}
						if (!"GA".equals(gagi3) && !"GI".equals(gagi3)) {
							if (Util.getIsEqueal(partmaterial3, MaterialNo)) {
								if ("��".equalsIgnoreCase(infolist.get(1))) {
									partmaterialFlag3 = false;
								}
							}
						}
					}
				}
				// �ų����������ĺ���
				if (partmaterialFlag1 && partmaterialFlag2 && partmaterialFlag3) {
					if (paramap.containsKey(weldno)) {
						String[] curenre = paramap.get(weldno);
						CurrentSerie = curenre[1];
						RecomWeldForce = curenre[0];
						CurrentSeriedfi = curenre[12];
					}
				}
				boolean flag = false;
				// �����1.2g��ǿ�ģ���׼�����ȡ���
				if (sheetstrength12.equals("1.2g")) {
					flag = true;
				}
				if (flag) {
					basethickness = getMinnum(vals[8], vals[12], vals[16]);
				}
				setStringCellAndStyle(sheet, importance, rowindex, 4, style, Cell.CELL_TYPE_STRING);
				setStringCellAndStyle(sheet, weldno, rowindex, 8, style6, Cell.CELL_TYPE_STRING);
				setStringCellAndStyle(sheet, boardnumber1, rowindex, 13, style, Cell.CELL_TYPE_STRING);
				setStringCellAndStyle(sheet, boardname1, rowindex, 16, style, Cell.CELL_TYPE_STRING);
				if (getIscontains1180(partmaterial1)) {
					XSSFCellStyle newstyle = getXSSFStyleByrgb(book, sheet, rowindex, 29, -1, new XSSFColor(new java.awt.Color(255,199,206)));
					setStringCellAndStyle2(sheet, partmaterial1, rowindex, 29, newstyle, Cell.CELL_TYPE_STRING);
				} else {
					XSSFCellStyle newstyle = getXSSFStyle(book, sheet,rowindex, 29, -1, IndexedColors.WHITE.getIndex());
					setStringCellAndStyle2(sheet, partmaterial1, rowindex, 29, newstyle, Cell.CELL_TYPE_STRING);
				}
//				if (!partmaterialFlag1) {
//					setStringCellAndStyle2(sheet, partmaterial1, rowindex, 29, stylepink, Cell.CELL_TYPE_STRING);
//				} else {
//					setStringCellAndStyle2(sheet, partmaterial1, rowindex, 29, style, Cell.CELL_TYPE_STRING);
//				}
				setStringCellAndStyle(sheet, partthickness1, rowindex, 36, style, 11);
				setStringCellAndStyle(sheet, boardnumber2, rowindex, 39, style, Cell.CELL_TYPE_STRING);
				setStringCellAndStyle(sheet, boardname2, rowindex, 42, style, Cell.CELL_TYPE_STRING);
				if (getIscontains1180(partmaterial2)) {
					XSSFCellStyle newstyle = getXSSFStyleByrgb(book, sheet, rowindex, 55, -1, new XSSFColor(new java.awt.Color(255,199,206)));
					setStringCellAndStyle2(sheet, partmaterial2, rowindex, 55, newstyle, Cell.CELL_TYPE_STRING);
				} else {
					XSSFCellStyle newstyle = getXSSFStyle(book, sheet,rowindex, 55, -1, IndexedColors.WHITE.getIndex());
					setStringCellAndStyle2(sheet, partmaterial2, rowindex, 55, newstyle, Cell.CELL_TYPE_STRING);
				}
//				if (!partmaterialFlag2) {
//					setStringCellAndStyle2(sheet, partmaterial2, rowindex, 55, stylepink, Cell.CELL_TYPE_STRING);
//				} else {
//					setStringCellAndStyle2(sheet, partmaterial2, rowindex, 55, style, Cell.CELL_TYPE_STRING);
//				}
				setStringCellAndStyle(sheet, partthickness2, rowindex, 62, style, 11);
				setStringCellAndStyle(sheet, boardnumber3, rowindex, 65, style, Cell.CELL_TYPE_STRING);
				setStringCellAndStyle(sheet, boardname3, rowindex, 68, style, Cell.CELL_TYPE_STRING);
				if (getIscontains1180(partmaterial3)) {
					XSSFCellStyle newstyle = getXSSFStyleByrgb(book, sheet, rowindex, 81, -1, new XSSFColor(new java.awt.Color(255,199,206)));
					setStringCellAndStyle2(sheet, partmaterial3, rowindex, 81, newstyle, Cell.CELL_TYPE_STRING);
				} else {
					XSSFCellStyle newstyle = getXSSFStyle(book, sheet,rowindex, 81, -1, IndexedColors.WHITE.getIndex());
					setStringCellAndStyle2(sheet, partmaterial3, rowindex, 81, newstyle, Cell.CELL_TYPE_STRING);
				}
//				if (!partmaterialFlag3) {
//					setStringCellAndStyle2(sheet, partmaterial3, rowindex, 81, stylepink, Cell.CELL_TYPE_STRING);
//				} else {
//					setStringCellAndStyle2(sheet, partmaterial3, rowindex, 81, style, Cell.CELL_TYPE_STRING);
//				}
				setStringCellAndStyle(sheet, partthickness3, rowindex, 87, style, 11);
				setStringCellAndStyle(sheet, layersnum, rowindex, 90, style, 10);
				setStringCellAndStyle(sheet, gagi, rowindex, 92, style, Cell.CELL_TYPE_STRING);
				setStringCellAndStyle(sheet, sheetstrength440, rowindex, 94, style, 10);
				setStringCellAndStyle(sheet, sheetstrength590, rowindex, 96, style, 10);
				setStringCellAndStyle(sheet, sheetstrength, rowindex, 98, style, 10);
				if (flag) {
					setStringCellAndStyle(sheet, "��", rowindex, 100, style, Cell.CELL_TYPE_STRING);
					if (getColorDistinction(layersnum, partmaterial1, partmaterial2, partmaterial3, partthickness1,
							partthickness2, partthickness3)) {
						XSSFCellStyle newstyle = getXSSFStyle(book, sheet, rowindex, 102, 1, IndexedColors.SKY_BLUE.getIndex());
						setStringCellAndStyle2(sheet, basethickness, rowindex, 102, newstyle, 11);
					} else {
						XSSFCellStyle newstyle = getXSSFStyle(book, sheet, rowindex, 102, 1, IndexedColors.VIOLET.getIndex());
						setStringCellAndStyle2(sheet, basethickness, rowindex, 102, newstyle, 11);
					}
					// �����������Ϊ��
					setStringCellAndStyle(sheet, "", rowindex, 105, style, Cell.CELL_TYPE_STRING);
					setStringCellAndStyle(sheet, "", rowindex, 108, style, Cell.CELL_TYPE_STRING);
					setStringCellAndStyle(sheet, "", rowindex, 111, style, Cell.CELL_TYPE_STRING);
				} else {
					XSSFCellStyle newstyle = getXSSFStyle(book, sheet, rowindex, 102, 12, IndexedColors.WHITE.getIndex());
					setStringCellAndStyle(sheet, basethickness, rowindex, 102, newstyle, 11);
					setStringCellAndStyle(sheet, "-", rowindex, 100, style, Cell.CELL_TYPE_STRING);
					setStringCellAndStyle(sheet, CurrentSerie, rowindex, 105, style, Cell.CELL_TYPE_STRING);
					setStringCellAndStyle(sheet, RecomWeldForce, rowindex, 108, style, Cell.CELL_TYPE_STRING);
					setStringCellAndStyle(sheet, CurrentSeriedfi, rowindex, 111, style, Cell.CELL_TYPE_STRING);
				}
			}
		}
	}
	private XSSFCellStyle getXSSFStyleByrgb(XSSFWorkbook book,XSSFSheet sheet,int rowindex,int cellindex,int colorindex,XSSFColor bgcolor)
	{
		XSSFRow row = sheet.getRow(rowindex);
		if(row!=null)
		{
			XSSFCell cell = row.getCell(cellindex);
			if(cell!=null)
			{
				XSSFCellStyle style = cell.getCellStyle();
				XSSFCellStyle newstyle = book.createCellStyle();
				newstyle = (XSSFCellStyle) style.clone();
				if(bgcolor != null)
				{
					newstyle.setFillForegroundColor(bgcolor);
					newstyle.setFillPattern(CellStyle.SOLID_FOREGROUND);
				}
				if(colorindex > -1)
				{
					// ����������ɫ
					Font font = book.createFont();
					Font sourcefont = style.getFont();
					font.setColor((short) colorindex);
					font.setFontHeightInPoints(sourcefont.getFontHeightInPoints());
					font.setFontName(sourcefont.getFontName());
					newstyle.setFont(font);
				}
			    return newstyle;
			}
		}
		return null;
	}

	/*
	 * �ж������1.2g��ǿ�ţ���Ҫ�ж�����/���Ǿ�������ɫ������Ϊ��ɫtrue������Ϊ��ɫfalse
	 */
	private boolean getColorDistinction(String layersnum, String partmaterial1, String partmaterial2,
			String partmaterial3, String partthickness1, String partthickness2, String partthickness3) {
		boolean flag = false;
		if (layersnum != null && !layersnum.isEmpty()) {
			int bznum = Integer.parseInt(layersnum);// �����
			if (bznum == 1) {
				flag = true;
			} else if (bznum == 2) { // ��������
				// ���ж��ǿ���1.2g��ǿ��
				boolean flag1 = getIscontains1180(partmaterial1);
				boolean flag2 = getIscontains1180(partmaterial2);
				boolean flag3 = getIscontains1180(partmaterial3);
				// ��һ����Ϊ�գ���Ҫ���������
				if (partmaterial1 == null || partmaterial1.isEmpty()) {
					// �������1.2g��ǿ��
					if (flag2 && flag3) {
						flag = false;
					} else {
						flag = getCompareresultByTwo(partmaterial2, partmaterial3, partthickness2, partthickness3,
								flag2, flag3);
					}
				} else if (partmaterial2 == null || partmaterial2.isEmpty()) {
					// �������1.2g��ǿ��
					if (flag1 && flag3) {
						flag = false;
					} else {
						flag = getCompareresultByTwo(partmaterial1, partmaterial3, partthickness1, partthickness3,
								flag1, flag3);
					}
				} else {
					// �������1.2g��ǿ��
					if (flag1 && flag2) {
						flag = false;
					} else {
						flag = getCompareresultByTwo(partmaterial1, partmaterial2, partthickness1, partthickness2,
								flag1, flag2);
					}
				}
			} else { // ��������
				// ���ж��ǿ���1.2g��ǿ��
				boolean flag1 = getIscontains1180(partmaterial1);
				boolean flag2 = getIscontains1180(partmaterial2);
				boolean flag3 = getIscontains1180(partmaterial3);
				// �����ǿ�ȶ���1.2G ������
				if (flag1 && flag2 && flag3) {
					flag = false;
				} else if (!flag1 && flag2 && flag3) { // ���1Ϊ��1.2g����������1.2g���
					// ��ȡ1.2g�еı���
					double h2 = getDoubleByString(partthickness2);
					double h3 = getDoubleByString(partthickness3);
					if (h2 < h3) {
						flag = getCompareresultByTwo(partmaterial1, partmaterial2, partthickness1, partthickness2,
								flag1, flag2);
					} else {
						flag = getCompareresultByTwo(partmaterial1, partmaterial3, partthickness1, partthickness3,
								flag1, flag3);
					}

				} else if (flag1 && !flag2 && flag3) { // ���2Ϊ��1.2g����������1.2g���
					// ��ȡ1.2g�еı���
					double h1 = getDoubleByString(partthickness1);
					double h3 = getDoubleByString(partthickness3);
					if (h1 < h3) {
						flag = getCompareresultByTwo(partmaterial2, partmaterial1, partthickness2, partthickness1,
								flag2, flag1);
					} else {
						flag = getCompareresultByTwo(partmaterial2, partmaterial3, partthickness2, partthickness3,
								flag2, flag3);
					}
				} else if (flag1 && flag2 && !flag3) { // ���3Ϊ��1.2g����������1.2g���
					// ��ȡ1.2g�еı���
					double h1 = getDoubleByString(partthickness1);
					double h2 = getDoubleByString(partthickness2);
					if (h1 < h2) {
						flag = getCompareresultByTwo(partmaterial3, partmaterial1, partthickness3, partthickness1,
								flag3, flag1);
					} else {
						flag = getCompareresultByTwo(partmaterial3, partmaterial2, partthickness3, partthickness2,
								flag3, flag2);
					}
				} else {// ֻ��һ��Ϊ1.2g��ǿ��
					double h1 = getDoubleByString(partthickness1);
					double h2 = getDoubleByString(partthickness2);
					double h3 = getDoubleByString(partthickness3);
					int kn1 = getSheetstrength(partmaterial1);
					int kn2 = getSheetstrength(partmaterial2);
					int kn3 = getSheetstrength(partmaterial3);

					if (h1 != h2 && h1 != h3 && h2 != h3) { // ������
						// 1.2G ����壬���ǣ������ȣ�
						if (flag1) {
							if (h1 < h2 && h1 < h3) {
								flag = false;
							} else { // 1.2G�����壬���Σ������ȣ� 1.2G�����У����Σ������ȣ�
								flag = true;
							}
						} else if (flag2) {
							if (h2 < h1 && h2 < h3) {
								flag = false;
							} else { // 1.2G�����壬���Σ������ȣ� 1.2G�����У����Σ������ȣ�
								flag = true;
							}
						} else {
							if (h3 < h1 && h3 < h2) {
								flag = false;
							} else { // 1.2G�����壬���Σ������ȣ� 1.2G�����У����Σ������ȣ�
								flag = true;
							}
						}
					} else { // ����1.2g��������������ͬ����Ƚ�ǿ�ȣ��������������ǿ�ȶ���1.2G�ߣ����ǣ��������������
						if (flag1) {
							if (kn1 < kn2 && kn1 < kn3) {
								flag = false;
							} else {
								flag = true;
							}
						} else if (flag2) {
							if (kn2 < kn1 && kn2 < kn3) {
								flag = false;
							} else {
								flag = true;
							}
						} else {
							if (kn3 < kn1 && kn3 < kn2) {
								flag = false;
							} else {
								flag = true;
							}
						}
					}
				}
			}
		}
		return flag;
	}

	/*
	 * �����ıȽ�
	 */
	private boolean getCompareresultByTwo(String partmaterial, String partmateria2, String partthickness1,
			String partthickness2, boolean flag1, boolean flag2) {
		boolean flag = false;
		// �жϰ���Ƿ���ͬ
		if (partthickness1.equals(partthickness2)) { // �����ͬ�����ж�ǿ��
			int kn1 = getSheetstrength(partmaterial);
			int kn2 = getSheetstrength(partmateria2);
			// ����İ�ǿ�ȶ���1.2G��ͣ�����
			if (flag1) {
				if (kn1 > kn2) {
					flag = true;
				} else {
					flag = false;
				}
			} else {
				if (kn1 > kn2) {
					flag = false;
				} else {
					flag = true;
				}
			}
		} else {// �����ͬ��1.2G ����壬����
			double high = 0.0;
			double ordinary = 0.0;
			if (flag1) {
				high = getDoubleByString(partthickness1);
				ordinary = getDoubleByString(partthickness2);
			} else {
				ordinary = getDoubleByString(partthickness1);
				high = getDoubleByString(partthickness2);
			}
			if (high > ordinary) {
				flag = true;
			} else {
				flag = false;
			}
		}
		return flag;
	}

	/*
	 * ���ݲ��ϻ�ȡǿ��
	 */
	private int getSheetstrength(String partmaterial) {
		int tkness = 0;
		if (partmaterial != null && !partmaterial.isEmpty()) {

			String Sheetstrength = "";
			String[] str = partmaterial.split("-");
			if (str.length > 1) {
				String tempstr = str[1].trim();
				if (tempstr != null && !"".equals(tempstr)) {
					for (int K = 0; K < tempstr.length(); K++) {
						if (tempstr.charAt(K) >= 48 && tempstr.charAt(K) <= 57) {
							Sheetstrength += tempstr.charAt(K);
						}
					}
				}
				if (!Sheetstrength.isEmpty()) {
					tkness = Integer.parseInt(Sheetstrength);
				}
			}
		}

		return tkness;
	}

	/*
	 * �ַ���תΪdouble�ͣ�Ϊ��Ĭ��Ϊ0.0
	 */
	private double getDoubleByString(String str) {
		double num = 0.0;
		if (str != null && !str.isEmpty()) {
			num = Double.parseDouble(str);
		}
		return num;
	}

	/*
	 * �жϲ����Ƿ���1180��Ҳ���Ǹ�ǿ��
	 */
	private boolean getIscontains1180(String partmaterial1) {
		boolean flag = false;
		if (partmaterial1 != null && !partmaterial1.isEmpty()) {
			String Sheetstrength = "";
			String[] str = partmaterial1.split("-");
			if (str.length > 1) {
				String tempstr = str[1].trim();
				if (tempstr != null && !"".equals(tempstr)) {
					for (int K = 0; K < tempstr.length(); K++) {
						if (tempstr.charAt(K) >= 48 && tempstr.charAt(K) <= 57) {
							Sheetstrength += tempstr.charAt(K);
						}
					}
				}
			}
			if (Sheetstrength.equals("1180")) {
				flag = true;
			}
		}
		return flag;
	}

	private void DealRSWQDSheet(XSSFWorkbook book, ReportViwePanel viewPanel) {
		// TODO Auto-generated method stub
		int sheetnum = 0;
		sheetnum = book.getNumberOfSheets();
		ArrayList sheetAtIndexs = new ArrayList(); // RSW��������λ��
		for (int i = 0; i < sheetnum; i++) {
			String sheetname = book.getSheetName(i);
			if (sheetname.contains("RSW����")) {
				sheetAtIndexs.add(i);
			}
		}
		if (sheetAtIndexs == null && sheetAtIndexs.size() < 1) {
			return;
		}
		// ��ȡsheet�ڵ�����
		ArrayList datalist = getSheetData(book, sheetAtIndexs, false);

		ArrayList hdlist = new ArrayList();

		// ���ݺ����ڻ�����Ϣ���л�ȡ�����Ϣ
		List hdinfo = new ArrayList();// ����������Ϣ

		hdinfo = getBoardInformation(baseinfolist, datalist);

		// ѭ��������Ϣ�����㲢��ȡ��������ֵ
		Map<String, String[]> paramap = getDealParameter(hdinfo, false, viewPanel);

		// �������������λ�ã�����ȡ������Ϣд�뵽������
		writeRSWQDInfomation(book, hdinfo, paramap);
	}

	private void writeRSWQDInfomation(XSSFWorkbook book, List hdinfo, Map<String, String[]> paramap) {
		// TODO Auto-generated method stub
		if (hdinfo != null && hdinfo.size() > 0) {
			// ����������ɫ
			Font font = book.createFont();
			font.setColor((short) 12);// ��ɫ����
			font.setFontHeightInPoints((short) 9);
			XSSFCellStyle style = book.createCellStyle();
			style.setBorderBottom(XSSFCellStyle.BORDER_THIN); // �±߿�
			style.setBorderLeft(XSSFCellStyle.BORDER_THIN);// ��߿�
			style.setBorderTop(XSSFCellStyle.BORDER_THIN);// �ϱ߿�
			style.setBorderRight(XSSFCellStyle.BORDER_THIN);// �ұ߿�
			style.setVerticalAlignment(XSSFCellStyle.VERTICAL_CENTER);
			style.setAlignment(XSSFCellStyle.ALIGN_CENTER);
			style.setFont(font);

			Font font2 = book.createFont();
			font2.setColor((short) 12);// ��ɫ����
			font2.setFontHeightInPoints((short) 18);
			font2.setBoldweight(Font.BOLDWEIGHT_BOLD); // ����Ӵ�
			XSSFCellStyle style2 = book.createCellStyle();
			style2.setBorderBottom(CellStyle.BORDER_MEDIUM); // ���߱߿�
			style2.setBorderLeft(CellStyle.BORDER_MEDIUM); // ���߱߿�
			style2.setBorderRight(CellStyle.BORDER_MEDIUM); // ���߱߿�
			style2.setBorderTop(CellStyle.BORDER_MEDIUM); // ���߱߿�
			style2.setVerticalAlignment(XSSFCellStyle.VERTICAL_CENTER);
			style2.setAlignment(XSSFCellStyle.ALIGN_CENTER);
			style2.setFont(font2);

			XSSFCellStyle style3 = book.createCellStyle();
			style3.setBorderBottom(XSSFCellStyle.BORDER_THIN); // �±߿�
			// style3.setBorderLeft(XSSFCellStyle.BORDER_THIN);// ��߿�
			style3.setBorderTop(XSSFCellStyle.BORDER_THIN);// �ϱ߿�
			style3.setBorderRight(XSSFCellStyle.BORDER_THIN);// �ұ߿�
			style3.setVerticalAlignment(XSSFCellStyle.VERTICAL_CENTER);
			style3.setAlignment(XSSFCellStyle.ALIGN_CENTER);
			style3.setFont(font);

			XSSFCellStyle style4 = book.createCellStyle();
			style4.setBorderBottom(XSSFCellStyle.BORDER_THIN); // �±߿�
			style4.setBorderLeft(XSSFCellStyle.BORDER_THIN);// ��߿�
			style4.setBorderTop(XSSFCellStyle.BORDER_THIN);// �ϱ߿�
			// style4.setBorderRight(XSSFCellStyle.BORDER_THIN);// �ұ߿�
			style4.setVerticalAlignment(XSSFCellStyle.VERTICAL_CENTER);
			style4.setAlignment(XSSFCellStyle.ALIGN_CENTER);
			style4.setFont(font);

			Font font3 = book.createFont();
			font3.setColor((short) 2);// ��ɫ����
			font3.setFontHeightInPoints((short) 10);
			XSSFCellStyle style33 = book.createCellStyle();
			style33.setBorderBottom(XSSFCellStyle.BORDER_THIN); // �±߿�
			// style3.setBorderLeft(XSSFCellStyle.BORDER_THIN);// ��߿�
			style33.setBorderTop(XSSFCellStyle.BORDER_THIN);// �ϱ߿�
			style33.setBorderRight(XSSFCellStyle.BORDER_THIN);// �ұ߿�
			style33.setVerticalAlignment(XSSFCellStyle.VERTICAL_CENTER);
			style33.setAlignment(XSSFCellStyle.ALIGN_CENTER);
			style33.setFont(font3);

			// ��ɫ����ɫ
			Font fontpink = book.createFont();
			fontpink.setColor((short) 12);// ��ɫ����
			fontpink.setFontName("MS PGothic");
			fontpink.setFontHeightInPoints((short) 9);

			XSSFCellStyle stylepink = book.createCellStyle();
			stylepink.setFillForegroundColor(IndexedColors.ROSE.getIndex());
			stylepink.setFillPattern(CellStyle.SOLID_FOREGROUND);
			stylepink.setBorderBottom(XSSFCellStyle.BORDER_THIN); // �±߿�
			stylepink.setBorderLeft(XSSFCellStyle.BORDER_THIN);// ��߿�
			stylepink.setBorderTop(XSSFCellStyle.BORDER_THIN);// �ϱ߿�
			stylepink.setBorderRight(XSSFCellStyle.BORDER_THIN);// �ұ߿�
			stylepink.setVerticalAlignment(XSSFCellStyle.VERTICAL_CENTER);
			stylepink.setAlignment(XSSFCellStyle.ALIGN_CENTER);
			stylepink.setFont(fontpink);

			for (int i = 0; i < hdinfo.size(); i++) {
				String[] vals = (String[]) hdinfo.get(i);
				int sheetindex = Integer.parseInt(vals[2]); // sheet����λ��
				int rowindex = Integer.parseInt(vals[1]); // ��������
				XSSFSheet sheet = book.getSheetAt(sheetindex);

				// ���ݲ��ʶ��ձ��жϺ����Ƿ������㺸�Ӳ���
				boolean partmaterialFlag1 = true;
				boolean partmaterialFlag2 = true;
				boolean partmaterialFlag3 = true;
				String partmaterial1 = vals[7];
				String partmaterial2 = vals[11];
				String partmaterial3 = vals[15];
				String gagi1 = vals[27];
				String gagi2 = vals[28];
				String gagi3 = vals[29];

				// ���ݲ��ʶ��ձ��ȡGA/GI����
				if (MaterialMap != null) {
					for (Map.Entry<String, List<String>> entry : MaterialMap.entrySet()) {
						String MaterialNo = entry.getKey();
						List<String> infolist = entry.getValue();
						if (!"GA".equals(gagi1) && !"GI".equals(gagi1)) {
							if (Util.getIsEqueal(partmaterial1, MaterialNo)) {
								if ("��".equalsIgnoreCase(infolist.get(1))) {
									partmaterialFlag1 = false;
								}
							}
						}
						if (!"GA".equals(gagi2) && !"GI".equals(gagi2)) {
							if (Util.getIsEqueal(partmaterial2, MaterialNo)) {
								if ("��".equalsIgnoreCase(infolist.get(1))) {
									partmaterialFlag2 = false;
								}
							}
						}
						if (!"GA".equals(gagi3) && !"GI".equals(gagi3)) {
							if (Util.getIsEqueal(partmaterial3, MaterialNo)) {
								if ("��".equalsIgnoreCase(infolist.get(1))) {
									partmaterialFlag3 = false;
								}
							}
						}
					}
				}

				setStringCellAndStyle(sheet, vals[4], rowindex, 4, style, Cell.CELL_TYPE_STRING);
				setStringCellAndStyle(sheet, vals[3], rowindex, 8, style3, Cell.CELL_TYPE_STRING);
				setStringCellAndStyle(sheet, vals[5], rowindex, 13, style, Cell.CELL_TYPE_STRING);
				setStringCellAndStyle(sheet, vals[6], rowindex, 16, style, Cell.CELL_TYPE_STRING);
//				if (!partmaterialFlag1) {
//					setStringCellAndStyle2(sheet, vals[7], rowindex, 29, stylepink, Cell.CELL_TYPE_STRING);
//				} else {
//					setStringCellAndStyle2(sheet, vals[7], rowindex, 29, style, Cell.CELL_TYPE_STRING);
//				}
				setStringCellAndStyle(sheet, vals[7], rowindex, 29, style, Cell.CELL_TYPE_STRING);
				
				setStringCellAndStyle(sheet, vals[8], rowindex, 36, style, 11);
				setStringCellAndStyle(sheet, vals[9], rowindex, 39, style, Cell.CELL_TYPE_STRING);
				setStringCellAndStyle(sheet, vals[10], rowindex, 42, style, Cell.CELL_TYPE_STRING);
//				if (!partmaterialFlag2) {
//					setStringCellAndStyle2(sheet, vals[11], rowindex, 55, stylepink, Cell.CELL_TYPE_STRING);
//				} else {
//					setStringCellAndStyle2(sheet, vals[11], rowindex, 55, style, Cell.CELL_TYPE_STRING);
//				}
				setStringCellAndStyle(sheet, vals[11], rowindex, 55, style, Cell.CELL_TYPE_STRING);
				
				setStringCellAndStyle(sheet, vals[12], rowindex, 62, style, 11);
				setStringCellAndStyle(sheet, vals[13], rowindex, 65, style, Cell.CELL_TYPE_STRING);
				setStringCellAndStyle(sheet, vals[14], rowindex, 68, style, Cell.CELL_TYPE_STRING);
//				if (!partmaterialFlag3) {
//					setStringCellAndStyle2(sheet, vals[15], rowindex, 81, stylepink, Cell.CELL_TYPE_STRING);
//				} else {
//					setStringCellAndStyle2(sheet, vals[15], rowindex, 81, style, Cell.CELL_TYPE_STRING);
//				}
				setStringCellAndStyle(sheet, vals[15], rowindex, 81, style, Cell.CELL_TYPE_STRING);
				
				setStringCellAndStyle(sheet, vals[16], rowindex, 88, style, 11);
				setStringCellAndStyle(sheet, vals[17], rowindex, 91, style, 10);
				setStringCellAndStyle(sheet, vals[18], rowindex, 93, style, Cell.CELL_TYPE_STRING);
				setStringCellAndStyle(sheet, vals[19], rowindex, 95, style, 10);
				setStringCellAndStyle(sheet, vals[20], rowindex, 97, style, 10);
				setStringCellAndStyle(sheet, vals[21], rowindex, 99, style, 10);
				setStringCellAndStyle(sheet, vals[22], rowindex, 102, style, 11);
				// �����1.2g��ǿ�ģ���׼�����ȡ���
				if (vals[23].equals("1.2g")) {
					setStringCellAndStyle(sheet, "", rowindex, 105, style, 10);
					setStringCellAndStyle(sheet, "", rowindex, 108, style, 10);
					setStringCellAndStyle(sheet, "", rowindex, 111, style, 10);
				} else {
					if (paramap.containsKey(vals[3])) {
						String[] paras = paramap.get(vals[3]);
						String poweroncurent2 = "";
						String CurrentSerie = "";
						String RecomWeldForce = "";
						// �ų����������ĺ���
						if (partmaterialFlag1 && partmaterialFlag2 && partmaterialFlag3) {
							poweroncurent2 = paras[7];
							CurrentSerie = paras[1];
							RecomWeldForce = paras[0];
						}
						setStringCellAndStyle(sheet, CurrentSerie, rowindex, 105, style, 10);
						setStringCellAndStyle(sheet, RecomWeldForce, rowindex, 108, style, 10);
						setStringCellAndStyle(sheet, CurrentSerie, rowindex, 111, style, 10);
					}
				}
			}
		}
	}

	private void DealPSWSheet(XSSFWorkbook book, ReportViwePanel viewPanel, String result) {
		// TODO Auto-generated method stub
		int sheetnum = 0;
		sheetnum = book.getNumberOfSheets();
		ArrayList sheetAtIndexs = new ArrayList(); // PSW����λ��
		for (int i = 0; i < sheetnum; i++) {
			String sheetname = book.getSheetName(i);
			if (sheetname.contains("PSW") && !sheetname.contains("�㺸")) {
				sheetAtIndexs.add(i);
			}
		}
		if (sheetAtIndexs == null && sheetAtIndexs.size() < 1) {
			return;
		}
		// ��ȡsheet�ڵ�����
		ArrayList datalist = getSheetData(book, sheetAtIndexs, false);

		ArrayList hdlist = new ArrayList();

		// ���ݺ����ڻ�����Ϣ���л�ȡ�����Ϣ
		List hdinfo = new ArrayList();// ����������Ϣ

		hdinfo = getBoardInformation(baseinfolist, datalist);

		// ѭ��������Ϣ�����㲢��ȡ��������ֵ
		Map<String, String[]> paramap = getDealParameter(hdinfo, false, viewPanel);

		// �������������λ�ã�����ȡ������Ϣд�뵽������
		writePSWInfomation(book, hdinfo, paramap);

		// ���»�ȡ�������ֵ
		ArrayList datalist2 = getSheetData(book, sheetAtIndexs, false);

		// ���ݺ�ǹ��ţ����¼��㺸ǹ�Ĳ���
		Map<String, String[]> Calculapara = getCalculapara(datalist2, paramap);

		System.out.println(Calculapara);
		// ����д�����õĺ��Ӳ���
		writePSWParaInfo(book, Calculapara, result);

	}

	private void writePSWParaInfo(XSSFWorkbook book, Map<String, String[]> calculapara, String result) {
		// TODO Auto-generated method stub
		if (calculapara != null && calculapara.size() > 0) {
			Font font2 = book.createFont();
			font2.setColor((short) 12);// ��ɫ����
			font2.setFontHeightInPoints((short) 11);
			font2.setBoldweight(Font.BOLDWEIGHT_BOLD); // ����Ӵ�
			XSSFCellStyle style2 = book.createCellStyle();
			style2.setBorderBottom(CellStyle.BORDER_MEDIUM); // ���߱߿�
			style2.setBorderLeft(CellStyle.BORDER_MEDIUM); // ���߱߿�
			style2.setBorderRight(CellStyle.BORDER_MEDIUM); // ���߱߿�
			style2.setBorderTop(CellStyle.BORDER_MEDIUM); // ���߱߿�
			style2.setVerticalAlignment(XSSFCellStyle.VERTICAL_CENTER);
			style2.setAlignment(XSSFCellStyle.ALIGN_CENTER);
			style2.setFont(font2);

			XSSFCellStyle style6 = book.createCellStyle();
			style6.setBorderBottom(XSSFCellStyle.BORDER_NONE); // �±߿�
			style6.setBorderLeft(XSSFCellStyle.BORDER_MEDIUM);// ��߿�
			style6.setBorderTop(XSSFCellStyle.BORDER_NONE);// �ϱ߿�
			style6.setBorderRight(XSSFCellStyle.BORDER_NONE);// �ұ߿�

			XSSFCellStyle style7 = book.createCellStyle();
			style7.setBorderBottom(XSSFCellStyle.BORDER_NONE); // �±߿�
			style7.setBorderLeft(XSSFCellStyle.BORDER_NONE);// ��߿�
			style7.setBorderTop(XSSFCellStyle.BORDER_NONE);// �ϱ߿�
			style7.setBorderRight(XSSFCellStyle.BORDER_NONE);// �ұ߿�

			for (Map.Entry<String, String[]> entry : calculapara.entrySet()) {
				String shindex = entry.getKey();
				String[] values = entry.getValue();
				if (Util.isNumber(shindex)) {
					int index = Integer.parseInt(shindex);
					XSSFSheet sheet = book.getSheetAt(index);
					XSSFRow terow = sheet.getRow(48);
					XSSFCell tecell = terow.getCell(108);
					String preedtion = tecell.getStringCellValue();
					boolean teflag = getIsSOPAfter(preedtion);
					System.out.println("�Ƿ�ΪSOP��" + teflag);
					if (!teflag) // д������
					{
						// ֻ��ѡ�����¼��㣬�Ż�д�����¼���Ĳ���ֵ
						if ("��".equals(result)) {
							setStringCellAndStyle(sheet, "��ѹ��", 5, 36, style2, Cell.CELL_TYPE_STRING);// ��ѹ��~����
							setStringCellAndStyle(sheet, "Ԥѹʱ��", 5, 42, style2, Cell.CELL_TYPE_STRING);// ��ѹ��~����
							setStringCellAndStyle(sheet, "����ʱ��", 5, 48, style2, Cell.CELL_TYPE_STRING);// ��ѹ��~����
							setStringCellAndStyle(sheet, "��һ          ͨ��ʱ��", 5, 54, style2, Cell.CELL_TYPE_STRING);// ��ѹ��~����
							setStringCellAndStyle(sheet, "��һ          ͨ�����", 5, 60, style2, Cell.CELL_TYPE_STRING);// ��ѹ��~����
							setStringCellAndStyle(sheet, "��ȴʱ��һ", 5, 66, style2, Cell.CELL_TYPE_STRING);// ��ѹ��~����
							setStringCellAndStyle(sheet, "�ڶ�          ͨ��ʱ��", 5, 72, style2, Cell.CELL_TYPE_STRING);// ��ѹ��~����
							setStringCellAndStyle(sheet, "�ڶ�          ͨ�����", 5, 78, style2, Cell.CELL_TYPE_STRING);// ��ѹ��~����
							setStringCellAndStyle(sheet, "��ȴʱ���", 5, 84, style2, Cell.CELL_TYPE_STRING);// ��ѹ��~����
							setStringCellAndStyle(sheet, "����          ͨ��ʱ��", 5, 90, style2, Cell.CELL_TYPE_STRING);// ��ѹ��~����
							setStringCellAndStyle(sheet, "����         ͨ�����", 5, 96, style2, Cell.CELL_TYPE_STRING);// ��ѹ��~����
							setStringCellAndStyle(sheet, "����", 5, 102, style2, Cell.CELL_TYPE_STRING);// ��ѹ��~����
							setStringCellAndStyle(sheet, "��ǯ�ѹ��", 5, 108, style2, Cell.CELL_TYPE_STRING);// ��ѹ��~����
																											// ElectrodeVol
							// �ٰѼ���Ĳ�����д��
							for (int j = 0; j < values.length; j++) {
								setStringCellAndStyle(sheet, values[j], 7, 36 + j * 6, style2, Cell.CELL_TYPE_STRING);// ��ѹ��~����
							}
						}
					} else {
						// ��պ��Ӳ���
						for (int j = 0; j < 3; j++) {
							for (int k = 0; k < 77; k++) {
								if (k == 0) {
									setStringCellAndStyle2(sheet, "", 5 + j, 36 + k, style6, Cell.CELL_TYPE_STRING);// ��ѹ��~����
								} else {
									setStringCellAndStyle2(sheet, "", 5 + j, 36 + k, style7, Cell.CELL_TYPE_STRING);// ��ѹ��~����
								}
							}
						}
					}

				}
			}
		}
	}

	private Map<String, String[]> getCalculapara(ArrayList datalist, Map<String, String[]> paramap) {
		// TODO Auto-generated method stub
		Map<String, String[]> map = new HashMap<String, String[]>();
		List tempguncode = new ArrayList();
		boolean sopflag = false;
		for (int i = 0; i < datalist.size(); i++) {
			String[] strVal = (String[]) datalist.get(i);
			String shindex = strVal[2];
			System.out.println("sheetλ�ã�" + shindex);
			if (shindex != null && !shindex.isEmpty()) { // �����ǹ�ı��Ϊ�գ������㣬��Ϊ�޷���֤׼ȷ��
				if (!tempguncode.contains(shindex)) {
					tempguncode.add(shindex);
				}
			}
		}
		System.out.println(tempguncode);

		if (tempguncode.size() > 0) {
			for (int i = 0; i < tempguncode.size(); i++) {
				boolean isCucalPara = true; // �Ƿ�����ۺϺ��Ӳ���
				int maxRepressure = 0;// ��ѹ�����ֵ
				int minRepressure = 99999999;// ��ѹ����Сֵ
				double sumrevalue = 0;// �ܵ���ֵ
				// List pages = new ArrayList();// ǹ��Ӧ��sheetҳ
				String guncode = (String) tempguncode.get(i);
				int nums = 0;
				for (int j = 0; j < datalist.size(); j++) {
					String[] values = (String[]) datalist.get(j);

					// ���ݲ��ʶ��ձ��жϺ����Ƿ������㺸�Ӳ���
					boolean partmaterialFlag1 = true;
					boolean partmaterialFlag2 = true;
					boolean partmaterialFlag3 = true;
					String partmaterial1 = "";
					String partmaterial2 = "";
					String partmaterial3 = "";
					if (guncode.equals(values[2])) {
						partmaterial1 = values[8];
						partmaterial2 = values[9];
						partmaterial3 = values[10];
					}

					// ���ݲ��ʶ��ձ��ȡGA/GI����
					if (MaterialMap != null) {
						for (Map.Entry<String, List<String>> entry : MaterialMap.entrySet()) {
							String MaterialNo = entry.getKey();
							List<String> infolist = entry.getValue();
							if (Util.getIsEqueal(partmaterial1, MaterialNo)) {
								if ("��".equalsIgnoreCase(infolist.get(1))) {
									partmaterialFlag1 = false;
									isCucalPara = false;
								}
							}

							if (Util.getIsEqueal(partmaterial2, MaterialNo)) {
								if ("��".equalsIgnoreCase(infolist.get(1))) {
									partmaterialFlag2 = false;
									isCucalPara = false;
								}

							}

							if (Util.getIsEqueal(partmaterial3, MaterialNo)) {
								if ("��".equalsIgnoreCase(infolist.get(1))) {
									partmaterialFlag3 = false;
									isCucalPara = false;
								}
							}

						}
					}
					// �ų����������ĺ���
					if (partmaterialFlag1 && partmaterialFlag2 && partmaterialFlag3) {
						if (guncode.equals(values[2])) {
							// ������λת��
							String temppower = values[6];
							System.out.println("��ȡ���ĵ���ֵ:" + temppower);
							if (Util.isNumber(temppower)) {
								double curent = (Double.parseDouble(temppower) / 1000);
								temppower = Double.toString(curent);
							}
							String poweroncurent2 = temppower;// �ڶ�ͨ�����
							String RecomWeldForce = values[5];// �Ƽ� ��ѹ��(N)
//							if (paramap.containsKey(values[3])) {
//								String[] Vals = paramap.get(values[3]);
//								poweroncurent2 = Vals[7];// �ڶ�ͨ�����
//								RecomWeldForce = Vals[0];// �Ƽ� ��ѹ��(N)
//							}
							if (Util.isNumber(RecomWeldForce)) {
								int repress = Integer.parseInt(RecomWeldForce);
								if (minRepressure > repress) {
									minRepressure = repress;
								}
								if (maxRepressure < repress) {
									maxRepressure = repress;
								}
							}
							if (Util.isNumber(poweroncurent2)) {
								sumrevalue = sumrevalue + Double.parseDouble(poweroncurent2);

								System.out.println("ƽ��ֵ�м䣺" + sumrevalue + poweroncurent2);
							}
							nums++;

						}
					}

//					if (!pages.contains(values[2])) {
//						pages.add(values[2]);
//					}
				}
				System.out.println("���ƽ��ֵ��" + sumrevalue);
				System.out.println("guncode��" + guncode);
				System.out.println("isCucalPara��" + isCucalPara);
				System.out.println("nums��" + nums);
				// �������
				String[] tatolcurenre = new String[12];
				if (!isCucalPara || nums == 0) {
					tatolcurenre[0] = "";
					tatolcurenre[1] = "";
					tatolcurenre[2] = "";
					tatolcurenre[3] = "";
					tatolcurenre[4] = "";
					tatolcurenre[5] = "";
					tatolcurenre[6] = "";
					tatolcurenre[7] = "";
					tatolcurenre[8] = "";
					tatolcurenre[9] = "";
					tatolcurenre[10] = "";
					tatolcurenre[11] = "";
				} else {
					// ���û����������Ͳ����㺸ǹ����
					tatolcurenre = getAverageParameterValues(cv, maxRepressure, minRepressure, sumrevalue, nums);
					if (minRepressure == 99999999) {
						tatolcurenre[0] = "";
					}
					if (sumrevalue == 0) {
						tatolcurenre[1] = "";
						tatolcurenre[2] = "";
						tatolcurenre[3] = "";
						tatolcurenre[4] = "";
						tatolcurenre[5] = "";
						tatolcurenre[6] = "";
						tatolcurenre[7] = "";
						tatolcurenre[8] = "";
						tatolcurenre[9] = "";
						tatolcurenre[10] = "";
						tatolcurenre[11] = "";
					}
				}

				map.put((String) guncode, tatolcurenre);
				System.out.println("��ȡ��MAp��" + map);
//				if (pages.size() > 0) {
//					for (int k = 0; k < pages.size(); k++) {
//						map.put((String) pages.get(k), tatolcurenre);
//					}
//				}
			}
		}
		return map;
	}

	// ����PSW��Ϣ
	private void writePSWInfomation(XSSFWorkbook book, List hdinfo, Map<String, String[]> paramap) {
		// TODO Auto-generated method stub
		if (hdinfo != null && hdinfo.size() > 0) {
			// ����������ɫ
			Font font = book.createFont();
			font.setColor((short) 12);// ��ɫ����
			font.setFontHeightInPoints((short) 9);
			XSSFCellStyle style = book.createCellStyle();
			style.setBorderBottom(XSSFCellStyle.BORDER_THIN); // �±߿�
			style.setBorderLeft(XSSFCellStyle.BORDER_THIN);// ��߿�
			style.setBorderTop(XSSFCellStyle.BORDER_THIN);// �ϱ߿�
			style.setBorderRight(XSSFCellStyle.BORDER_THIN);// �ұ߿�
			style.setVerticalAlignment(XSSFCellStyle.VERTICAL_CENTER);
			style.setAlignment(XSSFCellStyle.ALIGN_CENTER);
			style.setFont(font);

			Font font2 = book.createFont();
			font2.setColor((short) 12);// ��ɫ����
			font2.setFontHeightInPoints((short) 11);
			font2.setBoldweight(Font.BOLDWEIGHT_BOLD); // ����Ӵ�
			XSSFCellStyle style2 = book.createCellStyle();
			style2.setBorderBottom(CellStyle.BORDER_MEDIUM); // ���߱߿�
			style2.setBorderLeft(CellStyle.BORDER_MEDIUM); // ���߱߿�
			style2.setBorderRight(CellStyle.BORDER_MEDIUM); // ���߱߿�
			style2.setBorderTop(CellStyle.BORDER_MEDIUM); // ���߱߿�
			style2.setVerticalAlignment(XSSFCellStyle.VERTICAL_CENTER);
			style2.setAlignment(XSSFCellStyle.ALIGN_CENTER);
			style2.setFont(font2);

			Font font3 = book.createFont();
			font3.setColor((short) 12);// ��ɫ����
			font3.setFontHeightInPoints((short) 14);
			font3.setBoldweight(Font.BOLDWEIGHT_BOLD); // ����Ӵ�
			XSSFCellStyle style3 = book.createCellStyle();
			style3.setBorderBottom(CellStyle.BORDER_MEDIUM); // ���߱߿�
			style3.setBorderLeft(CellStyle.BORDER_MEDIUM); // ���߱߿�
			style3.setBorderRight(CellStyle.BORDER_MEDIUM); // ���߱߿�
			style3.setBorderTop(CellStyle.BORDER_MEDIUM); // ���߱߿�
			style3.setVerticalAlignment(XSSFCellStyle.VERTICAL_CENTER);
			style3.setAlignment(XSSFCellStyle.ALIGN_CENTER);
			style3.setFont(font3);

			XSSFCellStyle style4 = book.createCellStyle();
			style4.setBorderBottom(XSSFCellStyle.BORDER_THIN); // �±߿�
			// style4.setBorderLeft(XSSFCellStyle.BORDER_THIN);// ��߿�
			style4.setBorderTop(XSSFCellStyle.BORDER_THIN);// �ϱ߿�
			style4.setBorderRight(XSSFCellStyle.BORDER_THIN);// �ұ߿�
			style4.setVerticalAlignment(XSSFCellStyle.VERTICAL_CENTER);
			style4.setAlignment(XSSFCellStyle.ALIGN_CENTER);
			style4.setFont(font);

			XSSFCellStyle style5 = book.createCellStyle();
			style5.setBorderBottom(XSSFCellStyle.BORDER_THIN); // �±߿�
			style5.setBorderLeft(XSSFCellStyle.BORDER_THIN);// ��߿�
			style5.setBorderTop(XSSFCellStyle.BORDER_THIN);// �ϱ߿�
			// style5.setBorderRight(XSSFCellStyle.BORDER_THIN);// �ұ߿�
			style5.setVerticalAlignment(XSSFCellStyle.VERTICAL_CENTER);
			style5.setAlignment(XSSFCellStyle.ALIGN_CENTER);
			style5.setFont(font);

			// ����������ɫ
			Font font4 = book.createFont();
			font4.setColor((short) 2);// ��ɫ����
			font4.setFontHeightInPoints((short) 10);

			XSSFCellStyle style44 = book.createCellStyle();
			style44.setBorderBottom(XSSFCellStyle.BORDER_THIN); // �±߿�
			// style4.setBorderLeft(XSSFCellStyle.BORDER_THIN);// ��߿�
			style44.setBorderTop(XSSFCellStyle.BORDER_THIN);// �ϱ߿�
			style44.setBorderRight(XSSFCellStyle.BORDER_THIN);// �ұ߿�
			style44.setVerticalAlignment(XSSFCellStyle.VERTICAL_CENTER);
			style44.setAlignment(XSSFCellStyle.ALIGN_CENTER);
			style44.setFont(font4);

			// ��ɫ����ɫ
			Font fontpink = book.createFont();
			fontpink.setColor((short) 12);// ��ɫ����
			fontpink.setFontName("MS PGothic");
			fontpink.setFontHeightInPoints((short) 9);

			XSSFCellStyle stylepink = book.createCellStyle();
			stylepink.setFillForegroundColor(IndexedColors.ROSE.getIndex());
			stylepink.setFillPattern(CellStyle.SOLID_FOREGROUND);
			stylepink.setBorderBottom(XSSFCellStyle.BORDER_THIN); // �±߿�
			stylepink.setBorderLeft(XSSFCellStyle.BORDER_THIN);// ��߿�
			stylepink.setBorderTop(XSSFCellStyle.BORDER_THIN);// �ϱ߿�
			stylepink.setBorderRight(XSSFCellStyle.BORDER_THIN);// �ұ߿�
			stylepink.setVerticalAlignment(XSSFCellStyle.VERTICAL_CENTER);
			stylepink.setAlignment(XSSFCellStyle.ALIGN_CENTER);
			stylepink.setFont(fontpink);

			for (int i = 0; i < hdinfo.size(); i++) {
				String[] vals = (String[]) hdinfo.get(i);
				int sheetindex = Integer.parseInt(vals[2]); // sheet����λ��
				int rowindex = Integer.parseInt(vals[1]); // ��������
				XSSFSheet sheet = book.getSheetAt(sheetindex);

				// ���ݲ��ʶ��ձ��жϺ����Ƿ������㺸�Ӳ���
				boolean partmaterialFlag1 = true;
				boolean partmaterialFlag2 = true;
				boolean partmaterialFlag3 = true;
				String partmaterial1 = vals[7];
				String partmaterial2 = vals[11];
				String partmaterial3 = vals[15];
				String gagi1 = vals[27];
				String gagi2 = vals[28];
				String gagi3 = vals[29];

				// ���ݲ��ʶ��ձ��ȡGA/GI����
				if (MaterialMap != null) {
					for (Map.Entry<String, List<String>> entry : MaterialMap.entrySet()) {
						String MaterialNo = entry.getKey();
						List<String> infolist = entry.getValue();
						if (!"GA".equals(gagi1) && !"GI".equals(gagi1)) {
							if (Util.getIsEqueal(partmaterial1, MaterialNo)) {
								if ("��".equalsIgnoreCase(infolist.get(1))) {
									partmaterialFlag1 = false;
								}
							}
						}
						if (!"GA".equals(gagi2) && !"GI".equals(gagi2)) {
							if (Util.getIsEqueal(partmaterial2, MaterialNo)) {
								if ("��".equalsIgnoreCase(infolist.get(1))) {
									partmaterialFlag2 = false;
								}
							}
						}
						if (!"GA".equals(gagi3) && !"GI".equals(gagi3)) {
							if (Util.getIsEqueal(partmaterial3, MaterialNo)) {
								if ("��".equalsIgnoreCase(infolist.get(1))) {
									partmaterialFlag3 = false;
								}
							}
						}
					}
				}

				setStringCellAndStyle(sheet, vals[4], rowindex, 4, style, Cell.CELL_TYPE_STRING);
				setStringCellAndStyle(sheet, vals[3], rowindex, 8, style4, Cell.CELL_TYPE_STRING);
				setStringCellAndStyle(sheet, vals[5], rowindex, 13, style, Cell.CELL_TYPE_STRING);
				setStringCellAndStyle(sheet, vals[6], rowindex, 16, style, Cell.CELL_TYPE_STRING);
//				if (!partmaterialFlag1) {
//					setStringCellAndStyle2(sheet, vals[7], rowindex, 29, stylepink, Cell.CELL_TYPE_STRING);
//				} else {
//					setStringCellAndStyle2(sheet, vals[7], rowindex, 29, style, Cell.CELL_TYPE_STRING);
//				}
				setStringCellAndStyle(sheet, vals[7], rowindex, 29, style, Cell.CELL_TYPE_STRING);

				setStringCellAndStyle(sheet, vals[8], rowindex, 36, style, 11);
				setStringCellAndStyle(sheet, vals[9], rowindex, 39, style, Cell.CELL_TYPE_STRING);
				setStringCellAndStyle(sheet, vals[10], rowindex, 42, style, Cell.CELL_TYPE_STRING);
//				if (!partmaterialFlag2) {
//					setStringCellAndStyle2(sheet, vals[11], rowindex, 55, stylepink, Cell.CELL_TYPE_STRING);
//				} else {
//					setStringCellAndStyle2(sheet, vals[11], rowindex, 55, style, Cell.CELL_TYPE_STRING);
//				}
				setStringCellAndStyle(sheet, vals[11], rowindex, 55, style, Cell.CELL_TYPE_STRING);

				setStringCellAndStyle(sheet, vals[12], rowindex, 62, style, 11);
				setStringCellAndStyle(sheet, vals[13], rowindex, 65, style, Cell.CELL_TYPE_STRING);
				setStringCellAndStyle(sheet, vals[14], rowindex, 68, style, Cell.CELL_TYPE_STRING);
//				if (!partmaterialFlag3) {
//					setStringCellAndStyle2(sheet, vals[15], rowindex, 81, stylepink, Cell.CELL_TYPE_STRING);
//				} else {
//					setStringCellAndStyle2(sheet, vals[15], rowindex, 81, style, Cell.CELL_TYPE_STRING);
//				}
				setStringCellAndStyle(sheet, vals[15], rowindex, 81, style, Cell.CELL_TYPE_STRING);
				
				setStringCellAndStyle(sheet, vals[16], rowindex, 88, style, 11);
				setStringCellAndStyle(sheet, vals[17], rowindex, 91, style, 10);
				setStringCellAndStyle(sheet, vals[18], rowindex, 93, style, Cell.CELL_TYPE_STRING);
				setStringCellAndStyle(sheet, vals[19], rowindex, 95, style, 10);
				setStringCellAndStyle(sheet, vals[20], rowindex, 97, style, 10);
				setStringCellAndStyle(sheet, vals[21], rowindex, 99, style, 10);
				setStringCellAndStyle(sheet, vals[22], rowindex, 102, style, 11);

				// �����1.2g��ǿ�ģ���׼�����ȡ���
				if (vals[23].equals("1.2g")) {
					setStringCellAndStyle(sheet, "", rowindex, 105, style, 10);
					setStringCellAndStyle(sheet, "", rowindex, 108, style, 10);
					setStringCellAndStyle(sheet, "", rowindex, 111, style, 10);
				} else {
					if (paramap.containsKey(vals[3])) {
						String[] paras = paramap.get(vals[3]);
						String poweroncurent2 = "";
						String CurrentSerie = "";
						String RecomWeldForce = "";

						// �ų����������ĺ���
						if (partmaterialFlag1 && partmaterialFlag2 && partmaterialFlag3) {
							poweroncurent2 = paras[7];
							CurrentSerie = paras[1];
							RecomWeldForce = paras[0];
						}

						// ������λת��
						if (Util.isNumber(poweroncurent2)) {
							int curent = 0;
							curent = (int) (Double.parseDouble(poweroncurent2) * 1000);
							poweroncurent2 = Integer.toString(curent);
						}

						setStringCellAndStyle(sheet, CurrentSerie, rowindex, 105, style, 10);
						setStringCellAndStyle(sheet, RecomWeldForce, rowindex, 108, style, 10);
						setStringCellAndStyle(sheet, poweroncurent2, rowindex, 111, style, 10);
					}

				}
			}
		}

	}

	// �Ե�Ԫ��ֵ
	public static void setStringCellAndStyle(XSSFSheet sheet, String value, int rowIndex, int cellIndex,
			XSSFCellStyle Style, int celltype) {

		// �����������ַ��͵����� 10Ϊ���ͣ�11Ϊdouble��

		XSSFRow row = sheet.getRow(rowIndex);
		if (row == null)
			row = sheet.createRow(rowIndex);
		XSSFCell cell = row.getCell(cellIndex);
		if (cell == null)
			cell = row.createCell(cellIndex);
		// cell.setCellType(celltype);
		if (value == null || value.isEmpty()) {
			cell.setCellType(Cell.CELL_TYPE_BLANK);
		} else {
			if (celltype == Cell.CELL_TYPE_STRING) {
				cell.setCellType(Cell.CELL_TYPE_STRING);
				cell.setCellValue(value);
			} else if (celltype == 10) {
				cell.setCellType(Cell.CELL_TYPE_NUMERIC);
				cell.setCellValue(Integer.parseInt(value));
			} else if (celltype == 11) {
				cell.setCellType(Cell.CELL_TYPE_NUMERIC);
				cell.setCellValue(Double.parseDouble(value));
			} else {
				cell.setCellType(Cell.CELL_TYPE_STRING);
				cell.setCellValue(value);
			}
		}

		// cell.setCellStyle(Style);

	}

	// �Ե�Ԫ��ֵ
	public static void setStringCellAndStyle2(XSSFSheet sheet, String value, int rowIndex, int cellIndex,
			XSSFCellStyle Style, int celltype) {

		// �����������ַ��͵����� 10Ϊ���ͣ�11Ϊdouble��

		XSSFRow row = sheet.getRow(rowIndex);
		if (row == null)
			row = sheet.createRow(rowIndex);
		XSSFCell cell = row.getCell(cellIndex);
		if (cell == null)
			cell = row.createCell(cellIndex);
		// cell.setCellType(celltype);
		if (value == null || value.isEmpty()) {
			cell.setCellType(Cell.CELL_TYPE_BLANK);
		} else {
			if (celltype == Cell.CELL_TYPE_STRING) {
				cell.setCellType(Cell.CELL_TYPE_STRING);
				cell.setCellValue(value);
			} else if (celltype == 10) {
				cell.setCellType(Cell.CELL_TYPE_NUMERIC);
				cell.setCellValue(Integer.parseInt(value));
			} else if (celltype == 11) {
				cell.setCellType(Cell.CELL_TYPE_NUMERIC);
				cell.setCellValue(Double.parseDouble(value));
			} else {
				cell.setCellType(Cell.CELL_TYPE_STRING);
				cell.setCellValue(value);
			}
		}

		cell.setCellStyle(Style);

	}

	/*
	 * �������
	 */
	private Map<String, String[]> getDealParameter(List hdinfo, boolean flag, ReportViwePanel viewPanel) {
		// TODO Auto-generated method stub
		Map<String, String[]> paramap = new HashMap<String, String[]>();
		for (int j = 0; j < hdinfo.size(); j++) {
			String[] str = new String[13];
			String[] wbinfo = (String[]) hdinfo.get(j);
			String wbNO = wbinfo[3];
			String boradnum = wbinfo[17];// �����
			String basethickness = wbinfo[22];// ��׼���
			String sheetstrength1 = wbinfo[24];// ���1ǿ��
			String sheetstrength2 = wbinfo[25];// ���2ǿ��
			String sheetstrength3 = wbinfo[26];// ���3ǿ��
			String partthickness1 = wbinfo[8];// ���1���
			String partthickness2 = wbinfo[12];// ���2���
			String partthickness3 = wbinfo[16];// ���3���
			String gagi1 = wbinfo[27];// ���1GA/GI��
			String gagi2 = wbinfo[28];// ���2GA/GI��
			String gagi3 = wbinfo[29];// ���3GA/GI��

			// �Ƽ���ѹ������
			String Repressure = "";
			Repressure = getRepressure(basethickness, boradnum, sheetstrength1, sheetstrength2, sheetstrength3);
			str[0] = Repressure;
			// 24���к��������趨�� �������к�
			String parameterSerialNo24 = "";
			parameterSerialNo24 = getParameterSerialNo24(basethickness, boradnum, gagi1, gagi2, gagi3, sheetstrength1,
					sheetstrength2, sheetstrength3);

			str[1] = parameterSerialNo24;

			// 255���к��������趨�� �������к�
			// ��ȡ���Ȳ�
			double thicknessdifference = getThicknessDifference(partthickness1, partthickness2, partthickness3,
					boradnum);
			String parameterSerialNo255 = "";
			parameterSerialNo255 = getParameterSerialNo255(basethickness, boradnum, gagi1, gagi2, gagi3, sheetstrength1,
					sheetstrength2, sheetstrength3, thicknessdifference);
			// �����������У��ղ��� ��Ҫ�����������ŷ� ���߼�����
			if (flag) {
				// ֻ��RSW�ŷ������Ƽ����У���Ӧ��
				str[1] = parameterSerialNo255;
				// 255���ж��ձ�
				String SequenceComparison = "";
				SequenceComparison = getSequenceComparison(parameterSerialNo255);
				str[12] = SequenceComparison;
				str[2] = "";
				str[3] = "";
				str[4] = "";
				str[5] = "";
				str[6] = "";
				str[7] = "";
				str[8] = "";
				str[9] = "";
				str[10] = "";
				str[11] = "";
			} else {
				// ֻ��PSW��RSW������Ҫ�������ֵ
				// 24���к��������趨�� �Ƽ� ����ֵ
				String[] recommendedvalue = getRecommendedvalue(parameterSerialNo24);
				str[2] = recommendedvalue[0];
				str[3] = recommendedvalue[1];
				str[4] = recommendedvalue[2];
				str[5] = recommendedvalue[3];
				str[6] = recommendedvalue[4];
				str[7] = recommendedvalue[5];
				str[8] = recommendedvalue[6];
				str[9] = recommendedvalue[7];
				str[10] = recommendedvalue[8];
				str[11] = recommendedvalue[9];
				str[12] = "";
			}
			paramap.put(wbNO, str);

		}

		return paramap;
	}

	// 24���к��������趨�� �Ƽ� ����ֵ
	public static String[] getRecommendedvalue(String parameterSerialNo24) {
		// TODO Auto-generated method stub
		String[] recommendedvalue = new String[10];
		for (int i = 0; i < cv.size(); i++) {
			CurrentandVoltage cvotage = cv.get(i);
			String serialNo = cvotage.getSequenceNo();
			if (serialNo != null && serialNo.equals(parameterSerialNo24)) {
				recommendedvalue[0] = cvotage.getBvalue();// ����ʱ��
				recommendedvalue[1] = cvotage.getCvalue();// ��һ ͨ��ʱ��
				recommendedvalue[2] = cvotage.getEvalue();// ��һ ͨ�����
				recommendedvalue[3] = cvotage.getFvalue();// ��ȴʱ��һ
				recommendedvalue[4] = cvotage.getGvalue();// �ڶ�ͨ��ʱ��
				recommendedvalue[5] = cvotage.getIvalue();// �ڶ�ͨ�����
				recommendedvalue[6] = cvotage.getJvalue();// ��ȴʱ���
				recommendedvalue[7] = cvotage.getKvalue();// ���� ͨ��ʱ��
				recommendedvalue[8] = cvotage.getMvalue();// ���� ͨ�����
				recommendedvalue[9] = cvotage.getNvalue();// ����
				break;
			}
		}
		return recommendedvalue;
	}

	// 255���ж��ձ�
	public static String getSequenceComparison(String parameterSerialNo255) {
		// TODO Auto-generated method stub
		String SequenceComparison = "";
		for (int i = 0; i < sct.size(); i++) {
			SequenceComparisonTable sctable = sct.get(i);
			Map<String, String> map = sctable.getValues();
			if (map.containsKey("S" + parameterSerialNo255)) {
				String value = map.get("S" + parameterSerialNo255);
				if (value.trim().length() < 2) {
					value = "0" + value;
				}
				SequenceComparison = sctable.getParameterGroup() + "-" + value;
				break;
			}
		}

		return SequenceComparison;
	}

	// 255���к��������趨�� �������к�
	public static String getParameterSerialNo255(String basethickness, String boradnum, String gagi1, String gagi2,
			String gagi3, String sheetstrength1, String sheetstrength2, String sheetstrength3,
			double thicknessdifference) {
		// TODO Auto-generated method stub
		String parameterSerialNo255 = "";
		int lnum = 0; // �������
		int ganum = 0; // GA������
		int high = 0;// ��ǿ������
		if (gagi1.isEmpty()) {
			lnum++;
		}
		if (gagi2.isEmpty()) {
			lnum++;
		}
		if (gagi3.isEmpty()) {
			lnum++;
		}
		if (boradnum.equals("2")) {
			lnum--;
		}
		if (gagi1.equals("GA") || gagi1.equals("GI")) {
			ganum++;
		}
		if (gagi2.equals("GA") || gagi1.equals("GI")) {
			ganum++;
		}
		if (gagi3.equals("GA") || gagi1.equals("GI")) {
			ganum++;
		}
		if (!sheetstrength1.isEmpty()) {
			high++;
		}
		if (!sheetstrength2.isEmpty()) {
			high++;
		}
		if (!sheetstrength3.isEmpty()) {
			high++;
		}
		for (int i = 0; i < SFswc.size(); i++) {
			SFSequenceWeldingConditionList sfsw = SFswc.get(i);
			String value = sfsw.getBasethickness();
			if (Util.isNumber(value) && Util.isNumber(basethickness)) {
				if (Double.parseDouble(value) == Double.parseDouble(basethickness)) {
					if (boradnum.equals("2")) {
						if (lnum == 2 && high != 0) {
							parameterSerialNo255 = sfsw.getBvalue();
						}
						if (lnum == 2 && high == 0) {
							parameterSerialNo255 = sfsw.getCvalue();
						}
						if (ganum == 1 && thicknessdifference <= 2.4) {
							parameterSerialNo255 = sfsw.getDvalue();
						}
						if (ganum == 1 && thicknessdifference > 2.4) {
							parameterSerialNo255 = sfsw.getEvalue();
						}
						if (ganum == 2 && thicknessdifference <= 2.4) {
							parameterSerialNo255 = sfsw.getFvalue();
						}
						if (ganum == 2 && thicknessdifference > 2.4) {
							parameterSerialNo255 = sfsw.getGvalue();
						}
					}
					if (boradnum.equals("3")) {
						if (lnum == 3 && (high == 2 || high == 3)) {
							parameterSerialNo255 = sfsw.getHvalue();
						}
						if (lnum == 3 && (high == 0 || high == 1)) {
							parameterSerialNo255 = sfsw.getIvalue();
						}
						if (ganum == 1 && thicknessdifference <= 2.4) {
							parameterSerialNo255 = sfsw.getJvalue();
						}
						if (ganum == 1 && thicknessdifference > 2.4) {
							parameterSerialNo255 = sfsw.getKvalue();
						}
						if ((ganum == 2 || ganum == 3) && thicknessdifference <= 2.4) {
							parameterSerialNo255 = sfsw.getLvalue();
						}
						if ((ganum == 2 || ganum == 3) && thicknessdifference > 2.4) {
							parameterSerialNo255 = sfsw.getMvalue();
						}
					}
					break;
				}
			}

		}
		return parameterSerialNo255;
	}

	// ��ȡ���Ȳ�
	public static double getThicknessDifference(String partthickness1, String partthickness2, String partthickness3,
			String boradnum) {
		// TODO Auto-generated method stub
		double pk1;
		double pk2;
		double pk3;
		double thicknessdifference = 0;
		if (Util.isNumber(partthickness1)) {
			pk1 = Double.parseDouble(partthickness1);
		} else {
			pk1 = -1.0;
		}
		if (Util.isNumber(partthickness2)) {
			pk2 = Double.parseDouble(partthickness2);
		} else {
			pk2 = -1.0;
		}
		if (Util.isNumber(partthickness3)) {
			pk3 = Double.parseDouble(partthickness3);
		} else {
			pk3 = -1.0;
		}
		if (boradnum.equals("2")) {
			if (pk1 == -1.0) {
				if (pk2 < pk3) {
					thicknessdifference = pk3 / pk2;
				} else {
					thicknessdifference = pk2 / pk3;
				}
			}
			if (pk2 == -1.0) {
				if (pk1 < pk3) {
					thicknessdifference = pk3 / pk1;
				} else {
					thicknessdifference = pk1 / pk3;
				}
			}
			if (pk3 == -1.0) {
				if (pk1 < pk2) {
					thicknessdifference = pk2 / pk1;
				} else {
					thicknessdifference = pk1 / pk2;
				}
			}
		}
		if (boradnum.equals("3")) {
			String minstr = getMinnum(partthickness1, partthickness2, partthickness3);
			String maxstr = getMaxnum(partthickness1, partthickness2, partthickness3);
			double min = Double.parseDouble(minstr);
			double max = Double.parseDouble(maxstr);
			thicknessdifference = max / min;
		}
		BigDecimal bd = new BigDecimal(thicknessdifference);
		BigDecimal fact = bd.setScale(1, BigDecimal.ROUND_HALF_UP);
		thicknessdifference = fact.doubleValue();
		return thicknessdifference;
	}

	// 24���к��������趨�� �������к�
	public static String getParameterSerialNo24(String basethickness, String boradnum, String gagi1, String gagi2,
			String gagi3, String sheetstrength1, String sheetstrength2, String sheetstrength3) {
		// TODO Auto-generated method stub
		String parameterSerialNo24 = "";
		int lnum = 0; // �������
		int ginum = 0; // GI������
		int ganum = 0; // GA������
		int high = 0;// ��ǿ������

		if (gagi1.isEmpty()) {
			lnum++;
		}
		if (gagi2.isEmpty()) {
			lnum++;
		}
		if (gagi3.isEmpty()) {
			lnum++;
		}
		if (boradnum.equals("2")) {
			lnum--;
		}

		if (gagi1.equals("GA")) {
			ganum++;
		}
		if (gagi2.equals("GA")) {
			ganum++;
		}
		if (gagi3.equals("GA")) {
			ganum++;
		}
		if (gagi1.equals("GI")) {
			ginum++;
		}
		if (gagi2.equals("GI")) {
			ginum++;
		}
		if (gagi3.equals("GI")) {
			ginum++;
		}
		if (!sheetstrength1.isEmpty()) {
			high++;
		}
		if (!sheetstrength2.isEmpty()) {
			high++;
		}
		if (!sheetstrength3.isEmpty()) {
			high++;
		}
		for (int i = 0; i < swc.size(); i++) {
			SequenceWeldingConditionList swcl = swc.get(i);
			String thick = swcl.getBasethickness();
			if (Util.isNumber(thick) && Util.isNumber(basethickness)) {
				if (Double.parseDouble(thick) == Double.parseDouble(basethickness)) {
					// ��������м���GA������GI��ʱ����GA�ĵ���GI�������ǡ�

					if (boradnum.equals("2")) {
						if (lnum == 2 && high == 0) {
							parameterSerialNo24 = swcl.getBvalue();
						}
						if (lnum == 2 && high != 0) {
							parameterSerialNo24 = swcl.getCvalue();
						}
						if (lnum == 1 && ganum == 1 && high == 0) {
							parameterSerialNo24 = swcl.getDvalue();
						}
						if (lnum == 1 && ganum == 1 && high != 0) {
							parameterSerialNo24 = swcl.getEvalue();
						}
						if (lnum == 1 && ganum == 1 && high != 0) {
							parameterSerialNo24 = swcl.getEvalue();
						}
						if (ganum == 2 && high == 0) {
							parameterSerialNo24 = swcl.getFvalue();
						}
						if (ganum == 2 && high != 0) {
							parameterSerialNo24 = swcl.getGvalue();
						}
						if (ginum == 1 && ganum == 1 && high == 0) {
							parameterSerialNo24 = swcl.getHvalue();
						}
						if (ginum == 1 && ganum == 1 && high != 0) {
							parameterSerialNo24 = swcl.getIvalue();
						}
						if (ginum == 2 && high == 0) {
							parameterSerialNo24 = swcl.getJvalue();
						}
						if (ginum == 2 && high != 0) {
							parameterSerialNo24 = swcl.getKvalue();
						}
					}
					if (boradnum.equals("3")) {
						if (lnum == 3 && high == 0) {
							parameterSerialNo24 = swcl.getLvalue();
						}
						if (lnum == 3 && high != 0) {
							parameterSerialNo24 = swcl.getMvalue();
						}
						if (ganum == 1 && ginum == 0 && high == 0) {
							parameterSerialNo24 = swcl.getNvalue();
						}
						if (ganum == 1 && ginum == 0 && high != 0) {
							parameterSerialNo24 = swcl.getOvalue();
						}
						if (ganum == 2 && ginum == 0 && high == 0) {
							parameterSerialNo24 = swcl.getPvalue();
						}
						if (ganum == 2 && ginum == 0 && high != 0) {
							parameterSerialNo24 = swcl.getQvalue();
						}
						if (ganum == 3 && ginum == 0 && high == 0) {
							parameterSerialNo24 = swcl.getRvalue();
						}
						if (ganum == 3 && ginum == 0 && high != 0 && high != 3) {
							parameterSerialNo24 = swcl.getSvalue();
						}
						if (ganum == 3 && ginum == 0 && high == 3) {
							parameterSerialNo24 = swcl.getTvalue();
						}
						if (ganum == 0 && ginum == 1 && high == 0) {
							parameterSerialNo24 = swcl.getUvalue();
						}
						if (ganum == 0 && ginum == 1 && high != 0) {
							parameterSerialNo24 = swcl.getVvalue();
						}
						if (lnum == 1 && ganum != 2 && high == 0) {
							parameterSerialNo24 = swcl.getWvalue();
						}
						if (lnum == 1 && ganum != 2 && high != 0) {
							parameterSerialNo24 = swcl.getXvalue();
						}
						if (lnum == 0 && ganum != 3 && high == 0) {
							parameterSerialNo24 = swcl.getYvalue();
						}
						if (lnum == 0 && ganum != 3 && high != 0 && high != 3) {
							parameterSerialNo24 = swcl.getZvalue();
						}
						if (lnum == 0 && ganum != 3 && high == 3) {
							parameterSerialNo24 = swcl.getAAvalue();
						}
					}
					break;
				}
			}

		}

		return parameterSerialNo24;
	}

	// ��ȡ��ѹ��
	public static String getRepressure(String basethickness, String boradnum, String sheetstrength1,
			String sheetstrength2, String sheetstrength3) {
		// TODO Auto-generated method stub
		String repressure = "";
		String distinguish = ""; // ����
		int num1 = 0;// 440��������
		int num2 = 0;// 440
		int num3 = 0;// 590Mpa780Mpa980Mpa
		// �������1180ǿ�Ȱ�ģ�����������У�Ĭ��Ϊ��
		int shstrength1 = getInteger(sheetstrength1);
		int shstrength2 = getInteger(sheetstrength2);
		int shstrength3 = getInteger(sheetstrength3);
		if (shstrength1 == 1180 || shstrength2 == 1180 || shstrength3 == 1180) {
			repressure = "";
			return repressure;
		}
		// �������1350ǿ�Ȱ�ģ�����ΪI
		if (shstrength1 == 1350 || shstrength2 == 1350 || shstrength3 == 1350) {
			distinguish = "��";
		} else {
			if (sheetstrength1.isEmpty()) {
				num1++;
			}
			if (shstrength1 == 440) {
				num2++;
			}
			if (shstrength1 == 590 || shstrength1 == 780 || shstrength1 == 980) {
				num3++;
			}
			if (sheetstrength2.isEmpty()) {
				num1++;
			}
			if (shstrength2 == 440) {
				num2++;
			}
			if (shstrength2 == 590 || shstrength2 == 780 || shstrength2 == 980) {
				num3++;
			}
			if (sheetstrength3.isEmpty()) {
				num1++;
			}
			if (shstrength3 == 440) {
				num2++;
			}
			if (shstrength3 == 590 || shstrength3 == 780 || shstrength3 == 980) {
				num3++;
			}

			// �ȸ����������򣬻�ȡ����,�ٸ���3���
			if (boradnum.equals("2")) {
				if (num1 == 3) {
					distinguish = "��";
				}
				if (num3 == 0 && num2 == 1) {
					distinguish = "��";
				}
				if (num2 == 0 && num3 == 1) {
					distinguish = "��";
				}
				if (num2 == 2) {
					distinguish = "��";
				}
				if (num2 == 1 && num3 == 1) {
					distinguish = "��";
				}
				if (num3 == 2) {
					distinguish = "��";
				}
			} else if (boradnum.equals("3")) {
				if (num1 == 3) {
					distinguish = "��";
				}
				if (num1 == 2 && num2 == 1) {
					distinguish = "��";
				}
				if (num1 == 2 && num3 == 1) {
					distinguish = "��";
				}
				if (num1 == 1 && num2 == 2) {
					distinguish = "��";
				}
				if (num1 == 1 && num2 == 1 && num3 == 1) {
					distinguish = "��";
				}
				if (num1 == 1 && num3 == 2) {
					distinguish = "��";
				}
				if (num2 == 3) {
					distinguish = "��";
				}
				if (num2 == 2 && num3 == 1) {
					distinguish = "��";
				}
				if (num2 == 1 && num3 == 2) {
					distinguish = "��";
				}
				if (num3 == 3) {
					distinguish = "��";
				}
			} else {
				repressure = "";
				return repressure;
			}

		}
		for (int i = 0; i < rp.size(); i++) {
			RecommendedPressure repre = rp.get(i);
			String thickness = repre.getBasethickness();
			if (Util.isNumber(thickness) && Util.isNumber(basethickness)) {
				if (Double.parseDouble(thickness) == Double.parseDouble(basethickness)) {
					if (distinguish.equals("��")) {
						repressure = repre.getBvalue();
					}
					if (distinguish.equals("��")) {
						repressure = repre.getCvalue();
					}
					if (distinguish.equals("��")) {
						repressure = repre.getDvalue();
					}
					if (distinguish.equals("��")) {
						repressure = repre.getEvalue();
					}
					if (distinguish.equals("��")) {
						repressure = repre.getFvalue();
					}
					break;
				}
			}

		}

		return repressure;
	}

	// ��ȡsheet�ڵ�����
	private ArrayList getSheetData(XSSFWorkbook book, ArrayList sheetAtIndexs, boolean flag) {
		// TODO Auto-generated method stub

		ArrayList resultDataList = new ArrayList();
		for (int i = 0; i < sheetAtIndexs.size(); i++) {
			int sheetindex = (int) sheetAtIndexs.get(i);
			Sheet sheet = book.getSheetAt(sheetindex);
			// У��sheet�Ƿ�Ϸ�
			if (sheet == null) {
				return null;
			}
			// ��ȡ��һ������
			int firstRowNum = sheet.getFirstRowNum();
			Row firstRow = (Row) sheet.getRow(firstRowNum);
			if (null == firstRow) {
				logger.warning("����Excelʧ�ܣ��ڵ�һ��û�ж�ȡ���κ����ݣ�");
			}

			// �Ȼ�ȡ��ǹ�ͺ�
			String guncode = "";
			Row rowgun = (Row) sheet.getRow(5);
			if (rowgun != null) {
				Cell cell = rowgun.getCell(19);
				if (cell != null) {
					guncode = convertCellValueToString(cell);
				}
			}
			// ��ȡ���
			String edition = "";
			Row rowedition = (Row) sheet.getRow(48);
			if (rowedition != null) {
				Cell cell = rowedition.getCell(108);
				if (cell != null) {
					edition = convertCellValueToString(cell);
				}
			}

			// ����ÿһ�е����ݣ��������ݶ���
			int rowStart = 11;
			int rowEnd = 47;
			for (int rowNum = rowStart; rowNum < rowEnd; rowNum++) {
				Row row = (Row) sheet.getRow(rowNum);
				if (null == row) {
					continue;
				}
				if (flag) {
					String[] resultData = convertRowToData2(row);
					if (null == resultData) {
						logger.warning("�� " + row.getRowNum() + "�����ݲ��Ϸ����Ѻ��ԣ�");
						continue;
					}
					resultData[0] = guncode;
					resultData[1] = Integer.toString(rowNum);// ��������
					resultData[2] = Integer.toString(sheetindex);// ����sheetҳλ��
					resultData[7] = edition;// ���

					resultDataList.add(resultData);
				} else {
					String[] resultData = convertRowToData(row);
					if (null == resultData) {
						logger.warning("�� " + row.getRowNum() + "�����ݲ��Ϸ����Ѻ��ԣ�");
						continue;
					}
					resultData[0] = guncode;
					resultData[1] = Integer.toString(rowNum);// ��������
					resultData[2] = Integer.toString(sheetindex);// ����sheetҳλ��
					resultData[7] = edition;// ���

					resultDataList.add(resultData);
				}

			}
		}

		return resultDataList;
	}

	private String[] convertRowToData(Row row) {
		// TODO Auto-generated method stub
		String[] data = new String[11];
		Cell cell;
		// �����
		cell = row.getCell(8);
		String weldno = convertCellValueToString(cell);
		if (weldno == null || weldno.isEmpty()) {
			return null;
		}
		data[3] = weldno;

		// �����
		cell = row.getCell(91);
		String boradnum = convertCellValueToString(cell);
		data[4] = boradnum;
		// �Ƽ� ��ѹ��(N)
		cell = row.getCell(108);
		data[5] = convertCellValueToString(cell);
		// �Ƽ� ����ֵ(A)
		cell = row.getCell(111);
		data[6] = convertCellValueToString(cell);

		// ����1
		cell = row.getCell(29);
		data[8] = convertCellValueToString(cell);
		// ����2
		cell = row.getCell(55);
		data[9] = convertCellValueToString(cell);
		// ����3
		cell = row.getCell(81);
		data[10] = convertCellValueToString(cell);

		return data;
	}

	private String[] convertRowToData2(Row row) {
		// TODO Auto-generated method stub
		String[] data = new String[11];
		Cell cell;
		// �����
		cell = row.getCell(8);
		String weldno = convertCellValueToString(cell);
		if (weldno == null || weldno.isEmpty()) {
			return null;
		}
		data[3] = weldno;
		// �����
		cell = row.getCell(90);
		String boradnum = convertCellValueToString(cell);
		data[4] = boradnum;
		// �Ƽ� ��ѹ��(N)
		cell = row.getCell(108);
		data[5] = convertCellValueToString(cell);
		// �Ƽ� ����ֵ(A)
		cell = row.getCell(111);
		data[6] = convertCellValueToString(cell);

		// ����1
		cell = row.getCell(29);
		data[8] = convertCellValueToString(cell);
		// ����2
		cell = row.getCell(55);
		data[9] = convertCellValueToString(cell);
		// ����3
		cell = row.getCell(81);
		data[10] = convertCellValueToString(cell);

		return data;
	}

	private static String convertCellValueToString(Cell cell) {
		if (cell == null) {
			return null;
		}
		String returnValue = null;
		if (cell.getCellType() == Cell.CELL_TYPE_BLANK) {

		} else {
			switch (cell.getCellType()) {
			case Cell.CELL_TYPE_NUMERIC: // ����
				Double doubleValue = cell.getNumericCellValue();
				// ��ʽ����ѧ��������ȡһλ����
				DecimalFormat df = new DecimalFormat("0");
				returnValue = df.format(doubleValue);
				break;
			case Cell.CELL_TYPE_STRING: // �ַ���
				// cell.setCellType(Cell.CELL_TYPE_STRING);
				returnValue = cell.getStringCellValue();
				break;
			case Cell.CELL_TYPE_BOOLEAN: // ����
				Boolean booleanValue = cell.getBooleanCellValue();
				returnValue = booleanValue.toString();
				break;
			case Cell.CELL_TYPE_BLANK: // ��ֵ
				break;
			case Cell.CELL_TYPE_FORMULA: // ��ʽ
				returnValue = cell.getCellFormula();
				break;
			case Cell.CELL_TYPE_ERROR: // ����
				break;
			default:
				break;
			}
		}
		return returnValue;
	}

	// ����ƽ������
	private String[] getAverageParameterValues(List<CurrentandVoltage> cv, int maxRepressure, int minRepressure,
			double sumrevalue, int size) {
		// TODO Auto-generated method stub
		String[] values = new String[12];
		String press = "";// ��ѹ��
		String Preloadingtime = "15c.";// Ԥѹʱ�� Ĭ�����ֵ
		String uptime = "";// ����ʱ��
		String powerontime1 = "";// ��һ ͨ��ʱ��
		String poweroncurent1 = "";// ��һ ͨ�����
		String coolingtime1 = "";// ��ȴʱ��һ
		String powerontime2 = "";// �ڶ�ͨ��ʱ��
		String poweroncurent2 = "";// �ڶ�ͨ�����
		String coolingtime2 = "";// ��ȴʱ���
		String powerontime3 = "";// ���� ͨ��ʱ��
		String poweroncurent3 = "";// ���� ͨ�����
		String maintain = "";// ����

		int prepress = (maxRepressure + minRepressure) / 2;
		press = Integer.toString(prepress);
		// ����ƽ��ֵ
		BigDecimal biga1 = new BigDecimal(Double.toString(sumrevalue));
		BigDecimal bigsize = new BigDecimal(Double.toString(size));
		double average = biga1.divide(bigsize, 8, BigDecimal.ROUND_HALF_UP).doubleValue();

		System.out.println("ƽ��ֵ��" + average);
		// 255���к��������趨�� ������ѹ
		CurrentandVoltage currentandVoltage = getCurrentandVoltage(average, cv);

		System.out.println("���Դ�ӡ��" + currentandVoltage.getSequenceNo());

		// ��������
		System.out.println("��������Ϊ7.7��" + getCurrentandVoltage(7.7, cv).getSequenceNo());
//		System.out.println("��������Ϊ7.2��" + getCurrentandVoltage(7.2, cv).getSequenceNo());
//		System.out.println("��������Ϊ8.3��" + getCurrentandVoltage(8.3, cv).getSequenceNo());
//		System.out.println("��������Ϊ9.25��" + getCurrentandVoltage(9.25, cv).getSequenceNo());
//		System.out.println("��������Ϊ16.8��" + getCurrentandVoltage(16.8, cv).getSequenceNo());
//		System.out.println("��������Ϊ18��" + getCurrentandVoltage(18, cv).getSequenceNo());

		if (currentandVoltage != null) {
			uptime = currentandVoltage.getBvalue() + "c.";// ����ʱ��
			powerontime1 = currentandVoltage.getCvalue() + "c.";// ��һ ͨ��ʱ��
			poweroncurent1 = currentandVoltage.getEvalue() + "KA";// ��һ ͨ�����
			coolingtime1 = currentandVoltage.getFvalue() + "c.";// ��ȴʱ��һ
			powerontime2 = currentandVoltage.getGvalue() + "c.";// �ڶ�ͨ��ʱ��
			poweroncurent2 = currentandVoltage.getIvalue() + "KA";// �ڶ�ͨ�����
			coolingtime2 = currentandVoltage.getJvalue() + "c.";// ��ȴʱ���
			powerontime3 = currentandVoltage.getKvalue() + "c.";// ���� ͨ��ʱ��
			poweroncurent3 = currentandVoltage.getMvalue() + "KA";// ���� ͨ�����
			maintain = currentandVoltage.getNvalue() + "c.";// ����;
		}
		values[0] = press + "N";
		values[1] = Preloadingtime;
		values[2] = uptime;
		values[3] = powerontime1;
		values[4] = poweroncurent1;
		values[5] = coolingtime1;
		values[6] = powerontime2;
		values[7] = poweroncurent2;
		values[8] = coolingtime2;
		values[9] = powerontime3;
		values[10] = poweroncurent3;
		values[11] = maintain;

		return values;
	}

	// 255���к��������趨�� ������ѹ
	private CurrentandVoltage getCurrentandVoltage(double average, List<CurrentandVoltage> cv) {
		// TODO Auto-generated method stub
		int index = 0;
		double fact = 0;
		double yushu = average % 0.5;
		if (yushu > 0) {
			fact = average + 0.5 - average % 0.5;
		} else {
			fact = average;
		}
		if (fact < 7) {
			fact = 7;
		}
		if (fact > 17) {
			fact = 17;
		}
		CurrentandVoltage voltage = cv.get(0);
		double first = Double.parseDouble(voltage.getIvalue());
		double difference = Math.abs(fact - first);
		for (int i = 0; i < cv.size(); i++) {
			CurrentandVoltage vol = cv.get(i);
			double bvalue = Double.parseDouble(vol.getIvalue());
			double diff = Math.abs(fact - bvalue);
			if (diff < difference) {
				if (!vol.getSequenceNo().equals("3") && !vol.getSequenceNo().equals("4")
						&& !vol.getSequenceNo().equals("5")) {
					index = i;
					difference = diff;
				}
			}
		}
		CurrentandVoltage factvaltage = cv.get(index);

		return factvaltage;
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

	/*
	 * ���ݺ����ڻ�����Ϣ���л�ȡ�����Ϣ
	 */
	private List getBoardInformation(List<WeldPointBoardInformation> baseinfolist, ArrayList hdlist) {
		// TODO Auto-generated method stub
		List totalinfo = new ArrayList();
		if (baseinfolist != null) {
			for (int i = 0; i < hdlist.size(); i++) {
				String[] str = (String[]) hdlist.get(i);
				if (str[4] == null || str[4].isEmpty()) { // ���ݰ�����Ƿ�Ϊ�գ�ȷ���Ƿ���Ҫ���»�ȡ��������
					String weldno = str[3];
					for (int j = 0; j < baseinfolist.size(); j++) {
						WeldPointBoardInformation wpb = baseinfolist.get(j);
						if (wpb.getWeldno() != null && weldno != null && wpb.getWeldno().equals(weldno)) {
							String[] values = new String[30];
							values[0] = str[0];
							values[1] = str[1];
							values[2] = str[2];
							values[3] = wpb.getWeldno(); // ������
							values[4] = wpb.getImportance(); // ��Ҫ��
							values[5] = wpb.getBoardnumber1(); // ���1���
							values[6] = wpb.getBoardname1(); // ���1����
							values[7] = wpb.getPartmaterial1(); // ���1����
							values[8] = wpb.getPartthickness1(); // ���1���
							values[9] = wpb.getBoardnumber2(); // ���2���
							values[10] = wpb.getBoardname2(); // ���2����
							values[11] = wpb.getPartmaterial2(); // ���2����
							values[12] = wpb.getPartthickness2(); // ���2���
							values[13] = wpb.getBoardnumber3(); // ���3���
							values[14] = wpb.getBoardname3(); // ���3����
							values[15] = wpb.getPartmaterial3(); // ���3����
							values[16] = wpb.getPartthickness3(); // ���3���
							values[17] = wpb.getLayersnum(); // �����
							if (wpb.getGagi() != null && !wpb.getGagi().isEmpty()) {
								values[18] = wpb.getGagi(); // GA /GI
							} else {
								values[18] = "-"; // GA /GI
							}
							values[19] = wpb.getSheetstrength440(); // ����ǿ��(Mpa)440
							values[20] = wpb.getSheetstrength590(); // ����ǿ��(Mpa)590
							values[21] = wpb.getSheetstrength(); // ����ǿ��(Mpa)>590
							values[22] = wpb.getBasethickness(); // ��׼���
							values[23] = wpb.getSheetstrength12(); // ����ǿ��(Mpa)1.2G
							values[24] = wpb.getStrength1();// ���1ǿ��
							values[25] = wpb.getStrength2();// ���2ǿ��
							values[26] = wpb.getStrength3();// ���3ǿ��
							values[27] = wpb.getGagi1();// ���1GA/GI��
							values[28] = wpb.getGagi2();// ���2GA/GI��
							values[29] = wpb.getGagi3();// ���3GA/GI��
							totalinfo.add(values);
							break; // �ҵ�����������ѭ����ֱ�Ӳ�����һ������
						}
					}
				}
			}
		} else {
			System.out.println("��ȡ������Ϣʧ�ܣ�");
		}

		return totalinfo;
	}

	/*
	 * �ַ�ת��������
	 */
	public static int getInteger(String str) {
		int num = -1;
		if (Util.isNumber(str)) {
			num = (int) Double.parseDouble(str);
		}
		return num;
	}

	/*
	 * ȡ��Сֵ
	 */
	public static String getMinnum(String str1, String str2, String str3) {
		String minstr = "";
		if (str1 == null || str1.isEmpty()) {
			str1 = "9999";
		}
		if (str2 == null || str2.isEmpty()) {
			str2 = "9999";
		}
		if (str3 == null || str3.isEmpty()) {
			str3 = "9999";
		}
		if (Double.parseDouble(str1) > Double.parseDouble(str2)) {
			if (Double.parseDouble(str2) > Double.parseDouble(str3)) {
				minstr = str3;
			} else {
				minstr = str2;
			}
		} else {
			if (Double.parseDouble(str1) > Double.parseDouble(str3)) {
				minstr = str3;
			} else {
				minstr = str1;
			}
		}
//		if (minstr.equals("9999")) {
//			minstr = "";
//		}
		return minstr;
	}

	/*
	 * ȡ���ֵ
	 */
	public static String getMaxnum(String str1, String str2, String str3) {
		String maxstr = "";
		if (str1 == null || str1.isEmpty()) {
			str1 = "-1";
		}
		if (str2 == null || str2.isEmpty()) {
			str2 = "-1";
		}
		if (str3 == null || str3.isEmpty()) {
			str3 = "-1";
		}
		if (Double.parseDouble(str1) > Double.parseDouble(str2)) {
			if (Double.parseDouble(str1) > Double.parseDouble(str3)) {
				maxstr = str1;
			} else {
				maxstr = str3;
			}
		} else {
			if (Double.parseDouble(str2) > Double.parseDouble(str3)) {
				maxstr = str2;
			} else {
				maxstr = str3;
			}
		}
//		if (maxstr.equals("-1")) {
//			maxstr = "";
//		}
		return maxstr;
	}

	/*
	 * �жϰ���Ƿ�ΪSOP��
	 */
	private boolean getIsSOPAfter(String bc) {
		boolean flag = false;
		ArrayList edition = getEditionSizeRule();
		if (edition != null && edition.size() > 0) {
			if (edition.contains(bc)) {
				return false;
			}
		}
		if (bc != null) {
			if (bc.length() == 1) {
				char c = bc.charAt(0);
				if (c >= 'A' && c <= 'Z') {
					flag = true;
				}
			}
			if (bc.length() == 2) {
				char c = bc.charAt(0);
				char cc = bc.charAt(1);
				if (c >= 'A' && c <= 'Z' && cc >= 'A' && cc <= 'Z') {
					flag = true;
				}
			}
		}
		return flag;
	}

	// ��ѯ�����ѡ���ȡ�����Ϣ
	private ArrayList getEditionSizeRule() {
		ArrayList rule = new ArrayList();
		try {
			AbstractAIFUIApplication app = AIFUtility.getCurrentApplication();
			TCSession session = (TCSession) app.getSession();
			TCPreferenceService preferenceService = session.getPreferenceService();

			String str = preferenceService.getPreferenceDescription("DFL9_get_version_information");
			if (str != null) {
				String[] values = preferenceService.getStringValues("DFL9_get_version_information");
				if (values != null) {
					for (int i = 0; i < values.length; i++) {
						String value = values[i];
						rule.add(value);
					}
				}
			}
			return rule;
		} catch (Exception e) {
			e.printStackTrace();
		}
		return rule;
	}
	
	private XSSFCellStyle getXSSFStyle(XSSFWorkbook book,XSSFSheet sheet,int rowindex,int cellindex,int colorindex,int bgcolor)
	{
		XSSFRow row = sheet.getRow(rowindex);
		if(row!=null)
		{
			XSSFCell cell = row.getCell(cellindex);
			if(cell!=null)
			{
				XSSFCellStyle style = cell.getCellStyle();
				if(bgcolor > -1)
				{
					style.setFillForegroundColor((short)bgcolor);
					style.setFillPattern(CellStyle.SOLID_FOREGROUND);
				}
				if(colorindex > -1)
				{
					// ����������ɫ
					Font font = book.createFont();
					Font sourcefont = style.getFont();
					font.setColor((short) colorindex);
					font.setFontHeightInPoints(sourcefont.getFontHeightInPoints());
					font.setFontName(sourcefont.getFontName());
					style.setFont(font);
				}
			    return style;
			}
		}
		return null;
	}
}
