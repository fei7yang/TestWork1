package com.dfl.report.workschedule;

import java.awt.Color;
import java.awt.image.BufferedImage;
import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.math.BigDecimal;
import java.rmi.AccessException;
import java.text.DecimalFormat;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Collections;
import java.util.Comparator;
import java.util.Date;
import java.util.HashMap;
import java.util.Iterator;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import java.util.Map.Entry;
import java.util.logging.Logger;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import javax.imageio.ImageIO;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.PrintSetup;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.ShapeTypes;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFClientAnchor;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFDrawing;
import org.apache.poi.xssf.usermodel.XSSFFormulaEvaluator;
import org.apache.poi.xssf.usermodel.XSSFPrintSetup;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFSimpleShape;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.dfl.report.ExcelReader.CoverInfomation;
import com.dfl.report.ExcelReader.CurrentandVoltage;
import com.dfl.report.ExcelReader.SequenceWeldingConditionList;
import com.dfl.report.ExcelReader.WeldPointBoardInformation;
import com.dfl.report.ExcelReader.baseinfoExcelReader;
import com.dfl.report.util.FileUtil;
import com.dfl.report.util.GenerateReportInfo;
import com.dfl.report.util.NewOutputDataToExcel;
import com.dfl.report.util.ReportUtils;
import com.dfl.report.util.ReportViwePanel;
import com.dfl.report.util.Util;
import com.teamcenter.rac.aif.AbstractAIFUIApplication;
import com.teamcenter.rac.aif.kernel.AIFComponentContext;
import com.teamcenter.rac.aif.kernel.InterfaceAIFComponent;
import com.teamcenter.rac.aifrcp.AIFUtility;
import com.teamcenter.rac.cme.kernel.bvr.FlowUtil;
import com.teamcenter.rac.cme.kernel.mfg.IMfgFlow;
import com.teamcenter.rac.cme.kernel.mfg.IMfgNode;
import com.teamcenter.rac.kernel.TCComponent;
import com.teamcenter.rac.kernel.TCComponentBOMLine;
import com.teamcenter.rac.kernel.TCComponentDataset;
import com.teamcenter.rac.kernel.TCComponentFolder;
import com.teamcenter.rac.kernel.TCComponentFolderType;
import com.teamcenter.rac.kernel.TCComponentGroup;
import com.teamcenter.rac.kernel.TCComponentItem;
import com.teamcenter.rac.kernel.TCComponentItemRevision;
import com.teamcenter.rac.kernel.TCComponentTcFile;
import com.teamcenter.rac.kernel.TCComponentUser;
import com.teamcenter.rac.kernel.TCException;
import com.teamcenter.rac.kernel.TCPreferenceService;
import com.teamcenter.rac.kernel.TCProperty;
import com.teamcenter.rac.kernel.TCSession;
import com.teamcenter.rac.util.MessageBox;
import com.teamcenter.soa.exceptions.NotLoadedException;

public class StationInformationTableOp {

	private AbstractAIFUIApplication app;
	private static Logger logger = Logger.getLogger(baseinfoExcelReader.class.getName()); // ??????????
	private LinkedHashMap<String, String> map = new LinkedHashMap<String, String>();// sheet??????sheet????
	private ArrayList list = new ArrayList();// sheet??????
	SimpleDateFormat df2 = new SimpleDateFormat("yyyy.MM");// ????????????
	private Map<String, String> projVehMap;// ????????????????????familycode??????
	private String Edition;// ????
	private String topfoldername;
	private String model;// ????????
	private String nameNO;// ????????
	private boolean IsSameout;
	private String VehicleNo = "";// ????????
	private ArrayList partlist = new ArrayList();// ??????????
	private ArrayList tempPartlist = new ArrayList();
	private LinkedHashMap<String, String> fymap = new LinkedHashMap<String, String>();// ????????????????
	private ArrayList<TCComponentBOMLine> Discretelist = new ArrayList<>();// ??????????????
	private TCComponentBOMLine topbl = new TCComponentBOMLine();// ??????????????BOP
	private List<WeldPointBoardInformation> baseinfolist;// ????????????????
	List<TCComponentDataset> datasetList = new ArrayList<TCComponentDataset>();
	List<TCComponentItemRevision> revlist = new ArrayList<TCComponentItemRevision>();
	private TCSession session;
	private ArrayList Import = new ArrayList();
	private boolean updateflag = false; // ????????????
	private Map<String, String[]> notelist = new HashMap<String, String[]>();// ??????????????????????????????????
	private List pswlist = new ArrayList();// ??????????????????????????
	private List rswqdlist = new ArrayList();// ??????????????????????????
	private List rswsflist = new ArrayList();// ??????????????????????????
	private ArrayList deletelist;
	private List<CurrentandVoltage> cv;
	private Map<String, List<String>> MaterialMap;
	private String stlr = "";// ????????????????????????????????1??????2????

	public StationInformationTableOp(AbstractAIFUIApplication app, ArrayList list, LinkedHashMap<String, String> map,
			String edition, String model, String nameNO, String topfoldername, boolean isSameout,
			List<CurrentandVoltage> cv, List<WeldPointBoardInformation> baseinfolist,
			Map<String, List<String>> materialMap) throws TCException, AccessException {
		// TODO Auto-generated constructor stub
		this.app = app;
		this.map = map;
		this.list = list;
		this.Edition = edition;
		this.topfoldername = topfoldername;
		this.model = model;
		this.nameNO = nameNO;
		this.IsSameout = isSameout;
		this.cv = cv;
		this.baseinfolist = baseinfolist;
		this.MaterialMap = materialMap;
		session = (TCSession) app.getSession();
		initUI();
	}

	public StationInformationTableOp(AbstractAIFUIApplication app, String edition, String topfoldername,
			List<CurrentandVoltage> cv, List<WeldPointBoardInformation> baseinfolist,
			Map<String, List<String>> materialMap) throws AccessException, TCException {
		// TODO Auto-generated constructor stub
		this.app = app;
		this.Edition = edition;
		this.topfoldername = topfoldername;
		this.cv = cv;
		this.baseinfolist = baseinfolist;
		session = (TCSession) app.getSession();
		this.Edition = edition;
		this.updateflag = true;
		this.MaterialMap = materialMap;
		initUI();
	}

	private void initUI() throws TCException, AccessException {
		// TODO Auto-generated method stub

		InterfaceAIFComponent[] ifc = app.getTargetComponents();
		TCComponentBOMLine gwbl = (TCComponentBOMLine) ifc[0];
		TCComponentItemRevision gwrev = gwbl.getItemRevision();
		// ????????????????
		TCComponentBOMLine xubl = gwbl.parent().parent();
		String childrenFoldername = Util.getProperty(xubl, "bl_rev_object_name").replace("_", ".").replace("-", ".")
				.replace(" ", ".");
		try {
			topbl = gwbl.window().getTopBOMLine();
		} catch (TCException e1) {
			// TODO Auto-generated catch block
			e1.printStackTrace();
		}
		TCComponentGroup group = session.getGroup();
		// ????
		String groupname = group.getLocalizedFullName();
		// ??????
		String department = "";
		if (groupname != null
				&& (groupname.contains("??????????") || groupname.contains("simultaneous Engineering Section"))) {
			department = "H30";
		} else if (groupname != null
				&& (groupname.contains("??????????") || groupname.contains("Body Assembly Engineering Section"))) {
			department = "VE2";
		} else {
			department = "VE2";
		}
		// ???? ????-???? ??????
		projVehMap = ReportUtils.getDFL_Project_VehicleNo();
		String FamlilyCode = "";
		FamlilyCode = Util.getProperty(topbl, "bl_rev_project_ids");// ????????
		if (projVehMap.size() < 1) {
			VehicleNo = FamlilyCode;
		} else {
			VehicleNo = projVehMap.get(FamlilyCode);
			if (VehicleNo == null) {
				if (FamlilyCode != null) {
					VehicleNo = FamlilyCode;
				}
			}
		}
		// ???????????? ,??????????????????????#??????????????????+??????????????????????????
		String linename = Util.getProperty(gwbl.parent(), "bl_rev_object_name");
		String staname = Util.getProperty(gwbl, "bl_rev_object_name");
		String assyname = "";
		if (staname.length() > 1) {
			if (staname.substring(0, 1).equals("#")) {
				if(linename.endsWith("LH") || linename.endsWith("RH"))
				{
					assyname = linename.trim() + staname;
				}
				else
				{
					assyname = linename.trim() + " " + staname;
				}
				
			} else {
				assyname = staname;
			}
		} else {
			if(linename.endsWith("LH") || linename.endsWith("RH"))
			{
				assyname = linename.trim() + staname;
			}
			else
			{
				assyname = linename.trim() + " " + staname;
			}
		}

		// getPartsinformation(gwbl);
		// ????????
		TCComponentBOMLine ssgwbl = getSymmetryState(gwbl.parent(), staname);
		// ????????????????????
		GenerateReportInfo info = new GenerateReportInfo();
		// ????????
		String procName = "";
		boolean isupdatesame = false;
		// ????????????????????????????????????????????????????????????
		if (updateflag) {
			info.setExist(false);
			info.setIsgoon(true);
			info.setAction(""); //$NON-NLS-1$
			info.setMeDocument(null);
			info.setDFL9_process_type("H"); //$NON-NLS-1$
			info.setDFL9_process_file_type("AB"); // $NON-NLS-1$
			info.setmeDocumentName(procName);
			info.setFlag(true);
			info.setProject_ids(topbl.getItemRevision());

			try {
				info = ReportUtils.beforeGenerateReportAction(gwbl.getItemRevision(), info);
			} catch (TCException e) {
				e.printStackTrace();
				// EclipseUtils.info("Error : " + e.getMessage()); //$NON-NLS-1$
				return;
			}
			TCComponentItemRevision docRev = info.getMeDocument();
			String docrevname = Util.getProperty(docRev, "object_name");
			if (docrevname.contains("???u??") || docrevname.contains("???u??")) {
				isupdatesame = true;
			}
		}

		// ????????????????????????????????????
		if ((!IsSameout && !updateflag) || (updateflag && !isupdatesame)) {
			ssgwbl = null;
		}
		String LRassyname = "";
		if (ssgwbl != null) {
			String linename2 = Util.getProperty(ssgwbl.parent(), "bl_rev_object_name");
			String staname2 = Util.getProperty(ssgwbl, "bl_rev_object_name");
			if (staname2.length() > 1) {
				if (staname2.substring(0, 1).equals("#")) {
					LRassyname = linename2 + " " + staname2;
				} else {
					LRassyname = staname2;
				}
			} else {
				LRassyname = linename2 + " " + staname2;
			}
		}

		// ??????????????????????????????????????????????????????????????????????????
		// String stationname = Util.getProperty(gwrev, "b8_ChineseName");// ????????
		String stationname = "";
		if (Util.getIsMEProcStat(gwbl.parent())) {
			stationname = Util.getProperty(gwbl.parent().getItemRevision(), "b8_ChineseName")
					+ Util.getProperty(gwbl, "bl_rev_object_name");// ????????????
		} else {
			stationname = Util.getProperty(gwbl.parent().getItemRevision(), "b8_ChineseName");// ????????????
		}
		if (ssgwbl != null) {
			if (linename != null && linename.length() > 1
					&& linename.substring(linename.length() - 2, linename.length()).equals("LH")) {
				stationname = stationname.replace("??", "").replace("??", "") + " ???u??";
				stlr = "1";
			} else {
				stationname = stationname.replace("??", "").replace("??", "") + " ???u??";
				stlr = "2";
			}

		}

		if (!updateflag) {
			// procName = nameNO + "." + stationname.replace("#", "");
			procName = nameNO + "." + stationname;
		}
		if (!updateflag) {
			info.setExist(false);
			info.setIsgoon(true);
			info.setAction(""); //$NON-NLS-1$
			info.setMeDocument(null);
			info.setDFL9_process_type("H"); //$NON-NLS-1$
			info.setDFL9_process_file_type("AB"); // $NON-NLS-1$
			info.setmeDocumentName(procName);
			info.setFlag(true);
			info.setProject_ids(topbl.getItemRevision());

			try {
				info = ReportUtils.beforeGenerateReportAction(gwbl.getItemRevision(), info);
			} catch (TCException e) {
				e.printStackTrace();
				// EclipseUtils.info("Error : " + e.getMessage()); //$NON-NLS-1$
				return;
			}
		}

		System.out.println("The action is completed before the report operation is generated.");

		if (!info.isIsgoon()) {
			return;
		}
		InputStream inputStream = null;
		if (updateflag) {
			TCComponentItemRevision docmentRev = info.getMeDocument();
			procName = Util.getProperty(docmentRev, "object_name");
			inputStream = baseinfoExcelReader.getFileinbyreadExcel(docmentRev, "IMAN_specification", procName);

			if (inputStream == null) {
				MessageBox.post("??????" + procName + "????????????????" + procName + "????????", "????????", MessageBox.ERROR);
				return;
			}
		}
		// ??????????????????????????
		ReportViwePanel viewPanel = new ReportViwePanel("????????");
		viewPanel.setVisible(true);

		viewPanel.addInfomation("????????????...\n", 10, 100);

		// ????????????????
		if (updateflag) {

		} else {
			if (model.equals("????????????")) {
				inputStream = FileUtil.getTemplateFile("DFL_Template_EngineeringWorkListStation");
				if (inputStream == null) {
					viewPanel.addInfomation("??????????????????????????????????????????????????(????????DFL_Template_EngineeringWorkListStation)\\n",
							100, 100);
					return;
				}
			} else if (model.equals("VIN??????????")) {
				inputStream = FileUtil.getTemplateFile("DFL_Template_EngineeringWorkVINCarve");
				if (inputStream == null) {
					viewPanel.addInfomation("????????????????????????VIN????????????????????????(????????DFL_Template_EngineeringWorkVINCarve)\\n",
							100, 100);
					return;
				}
			} else {
				inputStream = FileUtil.getTemplateFile("DFL_Template_AdjustmentLine");
				if (inputStream == null) {
					viewPanel.addInfomation("????????????????????????????????????????????????(????????DFL_Template_AdjustmentLine)\\n", 100, 100);
					return;
				}
			}

			System.out.println("??????????????");
		}
		XSSFWorkbook book = null;
		if (updateflag) {
			book = NewOutputDataToExcel.creatXSSFWorkbook(inputStream);
		} else {
			// ??????????????sheet????????????????
			book = creatEngineeringXSSFWorkbook(inputStream, list, map);

			System.out.println("??????sheet??????");
		}
		viewPanel.addInfomation("????????????...\n", 20, 100);

		// ??????????????????
		ArrayList plist = new ArrayList();// ????????????????????????
		// String username = app.getSession().getUserName();// ??????
		TCComponentUser user = session.getUser();
		String username = user.getUserName();
		// ????????BOP??????????????????
		String objectname = Util.getProperty(topbl.getItemRevision(), "object_name");
		String factoryline = ReportUtils.getFactoryLineByBOP(objectname);
		String factory = "";
		String linebody = "";
		if (factoryline.length() > 3) {
			factory = factoryline.substring(0, 3);
			linebody = factoryline.substring(factoryline.length() - 1);
		}
		String baseCarType = VehicleNo + "-" + factory + "-NO" + linebody;
		String stationcode = "";
		String[] assynos;
		TCProperty p = gwrev.getTCProperty("b8_ProcAssyNo2");//b8_ProcAssyNo2
		if (p != null) {
			assynos = p.getStringValueArray(); // ???????? ????
		} else {
			assynos = null;
		}
		boolean rLflag = false;
		String[] assynos2 = null;
		if (ssgwbl != null) {
			TCProperty p2 = ssgwbl.getItemRevision().getTCProperty("b8_ProcAssyNo2");
			if (p2 != null) {
				assynos2 = p2.getStringValueArray(); // ???????? ????
			} else {
				assynos2 = null;
			}
			rLflag = true;
		}
		String LRsunffix = "";
		if (assynos2 != null && assynos2.length > 0) {
			if (assynos2[0] != null) {
				if (assynos2[0].length() >= 5) {
					LRsunffix = "/" + assynos2[0].trim().substring(4, 5);// ????????
				}
			}
		}
		if (assynos != null && assynos.length > 0) {
			if (assynos[0] != null) {
				if (assynos[0].length() >= 5) {
					stationcode = "M" + assynos[0].trim().substring(0, 5) + LRsunffix;// ????????
				} else {
					stationcode = "M" + assynos[0].trim() + LRsunffix;
				}
			}
		} else {
			stationcode = "";// ????????
		}

		// ????assy??
		List assylist = new ArrayList();
		List assynamelist = new ArrayList();
		if (assynos != null && assynos.length > 0) {
			for (int i = 0; i < assynos.length; i++) {
				if (assynos[i] != null) {
					String[] str = new String[2];
					str[0] = assynos[i];
					str[1] = assyname;
					assynamelist.add(str);
					assylist.add(assynos[i]);
				}
			}
		}
		if (assynos2 != null && assynos2.length > 0) {
			for (int i = 0; i < assynos2.length; i++) {
				if (assynos2[i] != null) {
					String[] str = new String[2];
					str[0] = assynos2[i];
					str[1] = LRassyname;
					assynamelist.add(str);
					assylist.add(assynos2[i]);
				}

			}
		}
		if (assynamelist != null) {
			for (int i = 0; i < assynamelist.size(); i++) {
				String[] str = (String[]) assynamelist.get(i);
				System.out.println("assynamelist????1??" + str[0]);
				System.out.println("assynamelist????2??" + str[1]);
			}
		}

		String pc = Edition;// ????
		plist.add(username);
		plist.add(df2.format(new Date()));// ????
		plist.add(baseCarType);
		plist.add(stationname);
		plist.add(stationcode);
		plist.add(pc);
		plist.add(department);

		// ????????????
		List RHlist = getNewPartsinformation(gwbl);
		List LHlist = new ArrayList();

		if (ssgwbl != null) {
			// RLflag = true;
			LHlist = getNewPartsinformation(ssgwbl);
		}
		// ??????????????
		SetLabelsAndSort(RHlist, gwbl, ssgwbl, LHlist);

		// getRLHStateData(sortList, LHlist);

		System.out.println("????????????????");

		viewPanel.addInfomation("", 30, 100);

		// ????????
		{
			Util.callByPass(session, true);
		}

		// ??????????????
		PartsinformationProcessing(book, assylist, assynamelist);

		System.out.println("??????????????????");

		// ????????????
		Map<String, File> piclist = getAll3DPictures(gwbl.getItemRevision());

		// ??????????????
		CompositionChartProcessing(book, assylist, assyname, rLflag, piclist);

		System.out.println("??????????????????");

		// ??????????????
		PoorPatternProcessing(book, assylist, rLflag);

		System.out.println("??????????????????");

		// ????????????????????
		Discretelist = Util.getChildrenByBOMLine(gwbl, "B8_BIWDiscreteOPRevision");

		List<TCComponentBOMLine> symmetryDiscretelist = new ArrayList<>(); // ????????????????????????
		if (ssgwbl != null) {
			// RLflag = true;
			symmetryDiscretelist = Util.getChildrenByBOMLine(ssgwbl, "B8_BIWDiscreteOPRevision");
		}

		System.out.println("????????????????????");

		viewPanel.addInfomation("", 40, 100);

		// ????????????????????????????????????????R??????????????PSW????RSWsheet??????????sheet????
		Map<String, TCComponentBOMLine> blmap = new LinkedHashMap<String, TCComponentBOMLine>();
		int psw = 0;
		int rswq = 0;
		int rsws = 0;
		List<String> Discretenamelist = new ArrayList<>(); // ??????????????????
		if (Discretelist.size() > 0) {
			for (int i = 0; i < Discretelist.size(); i++) {
				TCComponentBOMLine bl = (TCComponentBOMLine) Discretelist.get(i);
				String Discretename = Util.getProperty(bl, "bl_rev_object_name");
				if (!Discretenamelist.contains(Discretename)) {
					Discretenamelist.add(Discretename);
				}
				if (Discretename.length() > 1) {
					if (Discretename.substring(0, 1).equals("R")) {

						// ????RSW??????RSW????????????????????????????
						// ????????sheet????????sheet????
						String sheetname = CopySheet(book, "RSW????", rswq);
						if (sheetname != null) {
							blmap.put(sheetname, bl);
							rswq++;
						}
						String sheetname1 = CopySheet(book, "RSW????", rsws);
						if (sheetname1 != null) {
							blmap.put(sheetname1, bl);
							rsws++;
						}

					} else {
						// PSW????????
						String sheetname = CopySheet(book, "PSW", psw);
						if (sheetname != null) {
							blmap.put(sheetname, bl);
							psw++;
						}
					}
				} else {
					System.out.println("The spot welding process name is incorrect and will not be processed.");
				}
			}
			/*************************
			 * ??????????????????
			 */
			if (updateflag) {
				// ??????????????????????????
				getPageNumberManagement(book);

				RSWSFClearSheetContext(book, "RSW????");
				RSWQDClearSheetContext(book, "RSW????");
				PSWClearSheetContext(book, "PSW");
			}
			/******************************/

		}
		// ????????????????????????????map
		Map<String, TCComponentBOMLine> symmetrymap = new HashMap<>();
		// ??????????????????????
		if (symmetryDiscretelist.size() > 0) {
			for (int i = 0; i < symmetryDiscretelist.size(); i++) {
				TCComponentBOMLine bl = (TCComponentBOMLine) symmetryDiscretelist.get(i);
				String Discretename = Util.getProperty(bl, "bl_rev_object_name");

				if (Discretenamelist.contains(Discretename)) {
					symmetrymap.put(Discretename, bl);
					continue;
				}
				if (Discretename.length() > 1) {
					if (Discretename.substring(0, 1).equals("R")) {

						// ????RSW??????RSW????????????????????????????
						// ????????sheet????????sheet????
						String sheetname = CopySheet(book, "RSW????", rswq);
						if (sheetname != null) {
							blmap.put(sheetname, bl);
							rswq++;
						}
						String sheetname1 = CopySheet(book, "RSW????", rsws);
						if (sheetname1 != null) {
							blmap.put(sheetname1, bl);
							rsws++;
						}

					} else {
						// PSW????????
						String sheetname = CopySheet(book, "PSW", psw);
						if (sheetname != null) {
							blmap.put(sheetname, bl);
							psw++;
						}
					}
				} else {
					System.out.println("The spot welding process name is incorrect and will not be processed.");
				}
			}
		}
		if (blmap.size() > 0) {

			// ????????????
//			String baseName = "222.????????";
//			baseinfolist = getBaseinfomation(topbl, baseName);

			for (Map.Entry<String, TCComponentBOMLine> entry : blmap.entrySet()) {
				String shname = entry.getKey();
				TCComponentBOMLine bl = entry.getValue();
				// ????RSW??????RSW????????????????????????????

				// RSW????????????
				if (shname.contains("RSW????")) {
					RSWpneumaticinformationProcessing(book, bl, gwbl, shname, symmetrymap);

					System.out.println("RSW????????????????");
				} else if (shname.contains("RSW????")) {

					// RSW????????????
					RSWServoinformationProcessing(book, bl, gwbl, shname, symmetrymap);

					System.out.println("RSW????????????????");
				} else {
					// PSW????????
					String error = PSWinformationProcessing(book, bl, shname, symmetrymap);

					if (!error.isEmpty()) {
						viewPanel.dispose();
						MessageBox.post(error, "????????", MessageBox.ERROR);
						return;
					}

					System.out.println("PSW????????????");
				}
			}
		}

		// ????????sheet??????????

		ProcessingGlueIcon(book, gwbl);

		// ????????sheet??????????

		ProcessingInstallationIcon(book, gwbl);

		System.out.println("??????????????????????");

		// ??????????????sheet??????
		ProcessingStatistics(book, stationname, Discretelist, symmetryDiscretelist,ssgwbl);

		System.out.println("??????????????????????");

		viewPanel.addInfomation("", 50, 100);

		// ????????????????????
		// book = saveFileAndgetFile(book,filename);

		/* ??????????????????????sheet?????????????????????????????????????????????? */
		// ????????sheet??????????
		writePublicDataToSheet(book, plist);

		// ??????????????sheet
		deleteSheets(book);

		// ????????????
		writeRepatPublicDataToSheet(book);

		System.out.println("????????sheet??????????????????");

		if (!updateflag) {
			// ??????????????
			ValidPageProcessing(book);

			System.out.println("??????????????????");
		}

		// ??????????????????????sheet??????????????
		SetSheetRename(book);

		// ??????sheet????????????????????????????????????????????
		// setCellFormula(book);

		// ????????????
		int shs = book.getNumberOfSheets();
		String[] contents = new String[shs];
		for (int i = 0; i < shs; i++) {
			String sheetname = book.getSheetName(i);
			contents[i] = sheetname;
		}
		// ????????????????????????
		TCProperty ppp = gwrev.getTCProperty("b8_OperationContent");
		if (ppp != null) {
			ppp.setStringValueArray(contents);
		}
		String[] weldimport = new String[1];
		if (Import.size() > 0) {
			String tempimport = "";
			for (int i = 0; i < Import.size(); i++) {
				String tempn = (String) Import.get(i);
				tempimport = tempimport + tempn;
			}
			weldimport[0] = tempimport;
		} else {
			weldimport[0] = "C";
		}
		TCProperty ppp2 = gwrev.getTCProperty("b8_WPImptLevel");
		if (ppp2 != null) {
			ppp2.setStringValueArray(weldimport);
		}
		// ??????????????????????????????
		setPropertyValue(gwrev, "b8_OPNo", stationcode);
		setPropertyValue(gwrev, "b8_STName", stationname);
		setPropertyValue(gwrev, "b8_OpSheetNumber", Integer.toString(shs));

		gwrev.lock();
		gwrev.save();
		gwrev.unlock();

		System.out.println("sheet??????????????????????");

		// String filename = Util.getProperty(gwbl, "bl_rev_object_name") + "??????????";
		viewPanel.addInfomation("??????????????????????...\n", 60, 100);

		String filename = Util.formatString(procName);
		NewOutputDataToExcel.exportFile(book, filename);

		String fullFileName = FileUtil.getReportFileName(filename);
		System.out.println(fullFileName);
		TCComponentDataset ds = Util.createDataset(session, filename, fullFileName, "MSExcelX", "excel");
		if (ds != null) {
			datasetList.add(ds);
		}
		revlist.add(gwrev);
		try {
			TCComponentItem docunment = ReportUtils.afterGenerateReportAction(datasetList, revlist, info, procName, "",
					session);
			saveFileToFolder(docunment, topfoldername, childrenFoldername, procName);

			// ??????????????????
			TCProperty[] pdoc = docunment.getTCProperties(new String[] { "dfl9_vehiclePlant", "dfl9_processArea" });
			if (pdoc != null) {
				if (pdoc.length > 1 && pdoc[0] != null) {
					pdoc[0].setStringValue(topfoldername);
					pdoc[1].setStringValue(childrenFoldername);
					docunment.lock();
					docunment.save();
					docunment.unlock();
				}
			}
//			docunment.getLatestItemRevision().setProperty("dfl9_vehiclePlant", topfoldername);
//			docunment.getLatestItemRevision().setProperty("dfl9_processArea", childrenFoldername);

		} catch (TCException e) {
			e.printStackTrace();
			// EclipseUtils.info(Messages.FixtureOperation_15 + e.getMessage());
			return;
		}

		// saveFiles(filename, gwbl);
		viewPanel.addInfomation("", 80, 100);

		// ????????
		{
			Util.callByPass(session, false);
		}

		viewPanel.addInfomation("????????????????????????????????????????????????...\n", 100, 100);

	}

	/**
	 * ?????????????????? 20200727 hgq
	 * 
	 * @param book
	 * @param stationname
	 * @param discretelist2
	 * @param symmetryDiscretelist
	 * @param ssgwbl 
	 */
	private void ProcessingStatistics(XSSFWorkbook book, String stationname, ArrayList<TCComponentBOMLine> discretelist2,
			List<TCComponentBOMLine> symmetryDiscretelist, TCComponentBOMLine ssgwbl) {
		// TODO Auto-generated method stub
		// ??????????????sheet
		if (!updateflag) 
		{
			for(int i=0;i<deletelist.size();i++)
			{
				if(deletelist.get(i).toString().contains("??????????"))
				{
					return;
				}
			}			
		}		
		int sheetnum = 0;
		sheetnum = book.getNumberOfSheets();
		int sheetAtIndex1 = -1; // ??????????????????
		for (int i = 0; i < sheetnum; i++) {
			String sheetname = book.getSheetName(i);
			if (sheetname.contains("??????????")) {
				sheetAtIndex1 = i;
				break;
			}
		}
		if (sheetAtIndex1 == -1) {
			return;
		}
        //????????????
		if(ssgwbl!=null)
		{
			if (!updateflag) 
			{
				//??????????????????????????????????
				XSSFSheet newsheet = book.cloneSheet(sheetAtIndex1);
				book.setSheetOrder(newsheet.getSheetName(), sheetAtIndex1+1);
				String modelname = book.getSheetName(sheetAtIndex1);
				book.setSheetName(sheetAtIndex1, modelname + "-??");
				book.setSheetName(sheetAtIndex1+1, modelname + "-??");
			}						
			//????????????????????????????????????????????LH????????
			List<TCComponentBOMLine> LHlist = new ArrayList<>();
			List<TCComponentBOMLine> RHlist = new ArrayList<>();
			String ssgwname = Util.getProperty(ssgwbl, "bl_rev_object_name");
//			if(ssgwname.length()>2 && (ssgwname.substring(ssgwname.length()-3).contains("LH") || ssgwname.substring(ssgwname.length()-3).contains("??")))
//			{
			if("1".equals(stlr))
			{
				LHlist = discretelist2;
				RHlist = symmetryDiscretelist;
			}
			else
			{
				LHlist = symmetryDiscretelist;
				RHlist = discretelist2;
				
			}
			
			WriteManagementStatistics(book,LHlist,"??????????-??",stationname);
			WriteManagementStatistics(book,RHlist,"??????????-??",stationname);
		}
		else //??????????????
		{
			WriteManagementStatistics(book,discretelist2,"??????????",stationname);			
		}
	}

	private void WriteManagementStatistics(XSSFWorkbook book, List<TCComponentBOMLine> discretelist2, String sheettypename,
			String stationname) {
		// TODO Auto-generated method stub
		int sheetnum = 0;
		sheetnum = book.getNumberOfSheets();
		int sheetAtIndex1 = -1; // ??????????????????
		for (int i = 0; i < sheetnum; i++) {
			String sheetname = book.getSheetName(i);
			if (sheetname.contains(sheettypename)) {
				sheetAtIndex1 = i;
				break;
			}
		}
		if (sheetAtIndex1 == -1) {
			return;
		}
		
		int datanum = discretelist2.size();
		int page = datanum / 12 + 1;           
		// ??????????????????????sheet????????????????
		if (datanum % 12 == 0) {
			if (page > 1) {
				page = page - 1;
			}
		}
		// ????page????1????????????sheet??
		int index = sheetAtIndex1 + 1;			
		int shnum = 0;
		List<String> olddatalist = new ArrayList<>();
		if (updateflag) {
			
			//????????????????????
			olddatalist =  getStatisticsData(book,sheettypename);
			
			int number = 0;
			for (int i = 0; i < sheetnum; i++) {
				String sheetname = book.getSheetName(i);
				if (sheetname.contains(sheettypename)) {
					if (sheetAtIndex1 <= i) {
						number++;
					}
				}
			}
			index = index + page - 1;
			if (number < page) {
				if (page - number > 0) {
					for (int i = 0; i < page - number; i++) {
						XSSFSheet newsheet = book.cloneSheet(sheetAtIndex1);
						book.setSheetOrder(newsheet.getSheetName(), index);
						index++;
					}
				}
			}
		} else {
			if (page > 1) {
				for (int i = 1; i < page; i++) {
					XSSFSheet newsheet = book.cloneSheet(sheetAtIndex1);
					book.setSheetOrder(newsheet.getSheetName(), index);
					index++;
				}
			}
		}
		System.out.println("page: " + page);
		if (updateflag) 
		{
			//??????????????????????????
			List<String> newnamelist = new ArrayList<>();
			for(int i=0;i<discretelist2.size(); i++)
			{
				TCComponentBOMLine  direbl = (TCComponentBOMLine) discretelist2.get(i);
				String name = Util.getProperty(direbl, "bl_rev_object_name").replace("\\", "-");
				if(!olddatalist.contains(name))
				{
					newnamelist.add(name);
				}
			}
			for (int i = sheetAtIndex1; i < index; i++) 
			{
				XSSFSheet sheet = book.getSheetAt(i);	
				setStringCellAndStyle(sheet, stationname, 8, 21, null, Cell.CELL_TYPE_STRING);
				if (i == index - 1) 
				{
					for (int j = 0; j + 12 * shnum < olddatalist.size()+newnamelist.size(); j++) 
					{
						if(j + 12 * shnum>olddatalist.size()-1)
						{
							String name = newnamelist.get(j + 12 * shnum - olddatalist.size());
							setStringCellAndStyle(sheet, name, 8 + j*3, 41, null, Cell.CELL_TYPE_STRING);
						}					
					}			
				}
				else 
				{
					for (int j = 0; j + 12 * shnum < 12 + 12 * shnum; j++) 
					{
						if(j + 12 * shnum>olddatalist.size()-1)
						{
							String name = newnamelist.get(j + 12 * shnum - olddatalist.size());
							setStringCellAndStyle(sheet, name, 8 + j*3, 41, null, Cell.CELL_TYPE_STRING);
						}						
					}
				}				
				shnum++;
			}	
		}
		else
		{
			for (int i = sheetAtIndex1; i < index; i++) 
			{
				XSSFSheet sheet = book.getSheetAt(i);	
				setStringCellAndStyle(sheet, stationname, 8, 21, null, Cell.CELL_TYPE_STRING);
				if (i == index - 1) 
				{
					for (int j = 0; j + 12 * shnum < discretelist2.size(); j++) 
					{
						TCComponentBOMLine  direbl = (TCComponentBOMLine) discretelist2.get(j + 12 * shnum);
						String name = Util.getProperty(direbl, "bl_rev_object_name").replace("\\", "-");
						setStringCellAndStyle(sheet, name, 8 + j*3, 41, null, Cell.CELL_TYPE_STRING);
					}			
				}
				else 
				{
					for (int j = 0; j + 12 * shnum < 12 + 12 * shnum; j++) 
					{
						TCComponentBOMLine  direbl = (TCComponentBOMLine) discretelist2.get(j + 12 * shnum);
						String name = Util.getProperty(direbl, "bl_rev_object_name").replace("\\", "-");
						setStringCellAndStyle(sheet, name, 8 + j*3, 41, null, Cell.CELL_TYPE_STRING);
					}
				}				
				shnum++;
			}		
		}
			
	}

	private List<String> getStatisticsData(XSSFWorkbook book, String sheettypename) {
		// TODO Auto-generated method stub
		List<String> valuelist = new ArrayList<>();
		int sheetnum = book.getNumberOfSheets();
		for (int i = 0; i < sheetnum; i++) {
			String sheetname = book.getSheetName(i);
			if (sheetname.contains(sheettypename)) {
				XSSFSheet sheet = book.getSheetAt(i);
				// ??????9????42??
				int rowStart = 8;
				int rowEnd = 41;
				for (int rowNum = rowStart; rowNum < rowEnd; rowNum++) {
					Row row = (Row) sheet.getRow(rowNum);
					if (null == row) {
						continue;
					}
					Cell cell;
					// ????????
					cell = row.getCell(41);
					String name = convertCellValueToString(cell, Cell.CELL_TYPE_STRING);
					if(name!=null && !name.isEmpty())
					{
						valuelist.add(name);
					}
					rowNum = rowNum + 2;
				}
			}
		}
		return valuelist;
	}

	// ????????excel??????????????????
	public XSSFWorkbook saveFileAndgetFile(XSSFWorkbook book, String reportname) {
		try {
			String fullFileName = FileUtil.getReportFileName(reportname);
			File file = new File(fullFileName);
			if (file.exists()) {
				file.delete();
				file = new File(fullFileName);
			}

			FileOutputStream fOut = new FileOutputStream(file);
			try {
				book.write(fOut);
				fOut.flush();
				fOut.close();
			} catch (IOException e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			}

			XSSFWorkbook newbook = new XSSFWorkbook(fullFileName);

			System.out.println("????????????sheet??????" + newbook.getNumberOfSheets());

			// ????????????
			if (file.exists()) {
				file.delete();
			}

			return newbook;

			// ????excel
			// openFile(fullFileName);
		} catch (FileNotFoundException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		return null;
	}

	/*
	 * ??????????
	 */
	private void setPropertyValue(TCComponent tcc, String property, String value) throws TCException {
		TCProperty p = tcc.getTCProperty(property);
		if (p != null) {
			p.setStringValue(value);
		}
	}

	/*
	 * ??????????????????????????????????????????????partlist??????????????????????????????????????????????????????????????????
	 */
	private List getRLHStateData(List sortList, List lHlist,Map<String,String> partNoToNummap) {
		// TODO Auto-generated method stub
		if (sortList != null && sortList.size() > 0) {
			for (int i = 0; i < sortList.size(); i++) {
				if (lHlist != null && lHlist.size() > 0) {
					String[] values = (String[]) sortList.get(i);
					String partName = values[2];
					String partNo = values[1];
					tempPartlist.add(values);	
					for (Iterator<String[]> it = lHlist.iterator(); it.hasNext();) {
						String[] vals = it.next();
						String partName2 = vals[2];
						String partNo2 = vals[1];
						//??????????????????????
						if(!partNo.equals(partNo2))
						{
							if ((partName != null && partName.length() > 2)
									&& (partName2 != null && partName2.length() > 2)) {
								if (partName.substring(0, partName.length() - 2)
										.equals(partName2.substring(0, partName2.length() - 2))) {
									// vals[0] = "";
									vals[7] = values[7];
									vals[0] = values[0];
									tempPartlist.add(vals);
									it.remove();
								}
							}
						}
						else
						{
							String qty = vals[3];
							partNoToNummap.put(partNo, qty);						
							it.remove();
						}
					}
				} else {
					String[] values = (String[]) sortList.get(i);
					tempPartlist.add(values);
				}
			}
		}
		return lHlist;

	}
	/** 
     * ????java????????????????????.??0 
     * @param s 
     * @return  
     */  
    public static String subZeroAndDot(String s){  
        if(s.indexOf(".") > 0){  
            s = s.replaceAll("0+?$", "");//??????????0  
            s = s.replaceAll("[.]$", "");//????????????.??????  
        }  
        return s;  
    }  

	/*
	 * ??????????????
	 */
	private void SetLabelsAndSort(List list, TCComponentBOMLine gwbl, TCComponentBOMLine ssgwbl, List lHlist)
			throws AccessException, TCException {

		// ????????????????????????????
		// List oneList = new ArrayList();
		if (list == null) {
			return;
		}

		Comparator comparator = getComParatorBysequenceno();
		Collections.sort(list, comparator);

		int label = 0; // ????????
		int num = 1;// ??????????????????????
		int Occupynum = 0;// ??????????0????????????????
		String prePartno = "";// ??????????5??????
		String[] bh = { "A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S",
				"T", "U", "V", "W", "X", "Y", "Z", "AA", "AB", "AC", "AD", "AE", "AF", "AG", "AH", "AI", "AJ", "AK",
				"AL", "AM", "AN", "AO", "AP", "AQ", "AR", "AS", "AT", "AU", "AV", "AW", "AX", "AY", "AZ" };
		// ????????
		Map<String, String> tempmap = new HashMap<String, String>();
		ArrayList tempPartlist1 = new ArrayList();
		for (int i = 0; i < list.size(); i++) {
			String[] str = (String[]) list.get(i);
			if (str[1].toString().length() > 5 && (str[1].contains(" ") || str[1].contains(" "))) {
				prePartno = str[1].toString().substring(0, 5);
			} else {
				prePartno = str[1].toString();
			}
			String note = tempmap.get(prePartno);
			// ??????????5??????????????????
			if (note != null && !note.isEmpty()) {
				str[7] = note;
				int spno = 0;
				for (int j = 0; j < bh.length; j++) {
					if (bh[j].equals(note)) {
						spno = j + 1 - Occupynum;
					}
				}
				if (!str[0].equals("0")) {
					str[0] = Integer.toString(spno); // ????????????????
				}
				String strnum = fymap.get(note);
				int newnum = Integer.parseInt(strnum) + 1;
				fymap.put(note, Integer.toString(newnum));
			} else {
				if (label < 52) {
					str[7] = bh[label];
					if (!str[0].equals("0")) {
						str[0] = Integer.toString(label + 1 - Occupynum); // ????????????????
					} else {
						Occupynum++;
					}
				} else {
					str[7] = "";
					System.out.println("????????????????????????");
				}
				fymap.put(bh[label], "1");
				tempmap.put(prePartno, bh[label]);
				label++;
			}
			tempPartlist1.add(str);

		}

		// ??????????????????????????????????????????????tempPartlist????????????????????????????
		Map<String,String> partNoToNummap = new HashMap<>();//??????????????????????map
		List remainList = getRLHStateData(tempPartlist1, lHlist,partNoToNummap);

		if (remainList != null && remainList.size() > 0) {
			for (int i = 0; i < remainList.size(); i++) {
				String[] str = (String[]) remainList.get(i);
				if (str[1].toString().length() > 5 && (str[1].contains(" ") || str[1].contains(" "))) {
					prePartno = str[1].toString().substring(0, 5);
				} else {
					prePartno = str[1].toString();
				}
				String note = tempmap.get(prePartno);
				// ??????????5??????????????????
				if (note != null && !note.isEmpty()) {
					str[7] = note;
					int spno = 0;
					for (int j = 0; j < bh.length; j++) {
						if (bh[j].equals(note)) {
							spno = j + 1 - Occupynum;
						}
					}
					if (!str[0].equals("0")) {
						str[0] = Integer.toString(spno); // ????????????????
					}
					String strnum = fymap.get(note);
					int newnum = Integer.parseInt(strnum) + 1;
					fymap.put(note, Integer.toString(newnum));
				} else {
					if (label < 52) {
						str[7] = bh[label];
						if (!str[0].equals("0")) {
							str[0] = Integer.toString(label + 1 - Occupynum); // ????????????????
						} else {
							Occupynum++;
						}
					} else {
						str[7] = "";
						System.out.println("????????????????????????");
					}
					fymap.put(bh[label], "1");
					tempmap.put(prePartno, bh[label]);
					label++;
				}
				tempPartlist.add(str);
			}
		}

		// ??????????????????
		List LHlist = getLastStationPartList(gwbl);

		if (LHlist != null && LHlist.size() > 0) {
			for (int i = 0; i < LHlist.size(); i++) {
				String[] strVal = (String[]) LHlist.get(i);
				strVal[7] = bh[label];
				strVal[0] = Integer.toString(label + 1 - Occupynum); // ????????????????
				tempPartlist.add(strVal);
			}
			if (fymap.containsKey(bh[label])) {
				String numstr = fymap.get(bh[label]);
				int newnum = Integer.parseInt(numstr) + 1;
				fymap.put(bh[label], Integer.toString(newnum));
			} else {
				fymap.put(bh[label], "1");
			}
		}

		if (ssgwbl != null) {
			List RHlist = getLastStationPartList(ssgwbl);
			if (RHlist != null && RHlist.size() > 0) {
				for (int i = 0; i < RHlist.size(); i++) {
					String[] strVal = (String[]) RHlist.get(i);
					strVal[7] = bh[label];
					strVal[0] = Integer.toString(label + 1 - Occupynum); // ????????????????
					tempPartlist.add(strVal);
				}
				if (fymap.containsKey(bh[label])) {
					String numstr = fymap.get(bh[label]);
					int newnum = Integer.parseInt(numstr) + 1;
					fymap.put(bh[label], Integer.toString(newnum));
				} else {
					fymap.put(bh[label], "1");
				}
			}
		}
		// ????????????
		Comparator comparator2 = getComParatorBybh();
		Collections.sort(tempPartlist, comparator2);

		String firstNo = "";
		for (int i = 0; i < tempPartlist.size(); i++) {
			String[] value = (String[]) tempPartlist.get(i);
			String partNo = value[1];
			 //????????????????????????????????????/??????
			if(partNoToNummap.containsKey(partNo))
			{
				if("1".equals(stlr))
				{
					value[3] = partNoToNummap.get(partNo) + "/" + value[3];
				}
				else
				{
					value[3] = value[3] + "/" + partNoToNummap.get(partNo);
				}
			}
			
			if (i == 0) {
				firstNo = value[7];
				partlist.add(value);
			} else {
				if (!firstNo.equals(value[7].toString())) {
					String[] str = new String[9];
					partlist.add(str);
					partlist.add(value);
					firstNo = value[7];
				} else {
					partlist.add(value);
				}
			}
		}
		System.out.println(partlist);
		// return oneList;
	}

	/*
	 * ????????????
	 */
	private List getNewPartsinformation(TCComponentBOMLine gwbl) throws TCException, AccessException {
		// TODO Auto-generated method stub
		ArrayList install = new ArrayList();
		ArrayList templist = new ArrayList();
		// ??????????????????????????????
		install = Util.getChildrenByBOMLine(gwbl, "B8_BIWOperationRevision");

		System.out.println("??????????????????????" + install.size());

		for (int i = 0; i < install.size(); i++) {
			// ??????????????????????
			Map<String, String> partsource = getSizeRule();
			TCComponentBOMLine bl = (TCComponentBOMLine) install.get(i);
			ArrayList bflist = new ArrayList();
			bflist = Util.getChildrenByBOMLine(bl, "DFL9SolItmPartRevision");
			System.out.println("??????????????????" + bflist.size());
			for (int j = 0; j < bflist.size(); j++) {
				String[] info = new String[9];
				TCComponentBOMLine bfbl = (TCComponentBOMLine) bflist.get(j);
				info[0] = Util.getProperty(bfbl, "bl_sequence_no");// ????????
				if (info[0].isEmpty()) {
					info[0] = "0";
				}
				info[1] = Util.getProperty(bfbl, "bl_DFL9SolItmPartRevision_dfl9_part_no");// ????????
				// info[2] = Util.getProperty(bfbl, "bl_rev_object_name");// ????????
				info[2] = Util.getProperty(bfbl.getItemRevision(), "dfl9_CADObjectName");// ????????
				info[3] = Util.getProperty(bfbl, "bl_quantity");// ????
				if (info[3] == null || info[3].isEmpty()) {
					info[3] = "1";
				}
				String partresoles = "";
				String partresValue = "";
				TCProperty p = bfbl.getTCProperty("B8_BiwManualMU");
				if (p != null) {
					String lovindex = p.getStringValue();
					if (lovindex != null && !lovindex.isEmpty()) {
						if (partsource.containsKey(lovindex)) {
							partresoles = partsource.get(lovindex);
						}
						partresValue = lovindex;

					}
				}
				// partresoles = Util.getProperty(bfbl, "B8_NoteManualMark");// ???????? ??????
				if (partresoles == null || partresoles.isEmpty()) {
					TCProperty p2 = bfbl.getTCProperty("B8_NoteIsBiwTrUnit");
					if (p2 != null) {
						String lovindex = p2.getStringValue();
						if (lovindex != null && !lovindex.isEmpty()) {
							if (partsource.containsKey(lovindex)) {
								partresoles = partsource.get(lovindex);
							}
							partresValue = lovindex;
						}
					}
					// partresoles = Util.getProperty(bfbl, "B8_NoteIsBiwTrUnit");// ???????? ??????
				}
				info[6] = partresoles;
				info[8] = partresValue;
				System.out.println(" ????????:" + partresoles);
				if (partresValue.equals("Stamping")) {
					String thick = Util.getProperty(bfbl, "bl_DFL9SolItmPartRevision_dfl9PartThickness");// ????
					if (Util.isNumber(thick)) {
						Double th = Double.parseDouble(thick);
						info[4] = String.format("%.2f", th);
					} else {
						info[4] = thick;
					}
					info[5] = Util.getProperty(bfbl, "bl_DFL9SolItmPartRevision_dfl9PartMaterial");// ????
					System.out.println(" ????:" + Util.getProperty(bfbl, "bl_DFL9SolItmPartRevision_dfl9PartMaterial"));
				} else {
					info[4] = "";// ????
					info[5] = "";// ????
				}
				templist.add(info);
			}
		}
		// ????????????????????????????????????
		Map<String, String[]> map = new HashMap<String, String[]>();
		for (int i = 0; i < templist.size(); i++) {
			String[] value = (String[]) templist.get(i);
			String key = value[1];
			if (!map.containsKey(key)) {
				map.put(key, value);
			} else {
				String[] oldstr = map.get(key);
				int quality = 0;
				quality = Integer.parseInt(oldstr[3]) + Integer.parseInt(value[3]);
				oldstr[3] = Integer.toString(quality);
				map.put(key, oldstr);
			}
		}
		List newtemplist = new ArrayList();
		for (Map.Entry<String, String[]> entry : map.entrySet()) {
			String[] values = entry.getValue();
			newtemplist.add(values);
		}
		return newtemplist;

	}

	/*
	 * ????????????????????assy????
	 */
	private List getLastStationPartList(TCComponentBOMLine bl) throws TCException, AccessException {
		List templist = new ArrayList();

		// ????????????????????????????????assy????
		TCProperty pp = bl.getTCProperty("Mfg0predecessors");// ????????
		// TCProperty pp = bl.getTCProperty("Mfg0successors");//????????
		if (pp != null) {
			TCComponent[] obj = pp.getReferenceValueArray();
			for (int i = 0; i < obj.length; i++) {
				TCComponentBOMLine prebl = (TCComponentBOMLine) obj[i];
				String sequence_no = Util.getProperty(prebl, "bl_sequence_no");// ????????
				if (sequence_no.isEmpty()) {
					sequence_no = "0";
				}
				String quantity = Util.getProperty(prebl, "bl_quantity");// ????
				if (quantity == null || quantity.isEmpty()) {
					quantity = "1";
				}
				// ???????????? ,??????????????????
				String linename = Util.getProperty(prebl.parent(), "bl_rev_object_name");
				String assyname = linename;

				TCProperty p = prebl.getItemRevision().getTCProperty("b8_ProcAssyNo2");
				String[] assynos;
				if (p != null) {
					assynos = p.getStringValueArray();// ???????? ????
				} else {
					assynos = null;
				}
				if (assynos != null && assynos.length > 0) {
					for (int j = 0; j < assynos.length; j++) {
						String[] info = new String[9];
						info[0] = sequence_no;// ????????
						info[1] = assynos[j];// ????????
						info[2] = assyname;// ????????
						info[3] = quantity;// ????
						info[4] = "";// ????
						info[5] = "";// ????
						info[6] = "????????";// ???????? ??????
						if(assynos[j] != null)
						{
							templist.add(info);
						}						
					}
				}
			}
		}
		// ????????????????????????????assy????
		// List<IMfgFlow> list = FlowUtil.getScopeOutputFlows(bl);//????????????
		List<IMfgFlow> list = FlowUtil.getScopeInputFlows(bl);// ????????????
		if (list != null && list.size() > 0) {
			for (IMfgFlow flow : list) {
				IMfgNode node = flow.getPredecessor();
				TCComponentBOMLine preComp = (TCComponentBOMLine) node.getComponent();
				String sequence_no = Util.getProperty(preComp, "bl_sequence_no");// ????????
				if (sequence_no.isEmpty()) {
					sequence_no = "0";
				}
				String quantity = Util.getProperty(preComp, "bl_quantity");// ????
				if (quantity == null || quantity.isEmpty()) {
					quantity = "1";
				}
				// ???????????? ,??????????????????
				String linename = Util.getProperty(preComp.parent(), "bl_rev_object_name");
				String assyname = linename;

				TCProperty p = preComp.getItemRevision().getTCProperty("b8_ProcAssyNo2");
				String[] assynos;
				if (p != null) {
					assynos = p.getStringValueArray();// ???????? ????
				} else {
					assynos = null;
				}
				if (assynos != null && assynos.length > 0) {
					for (int j = 0; j < assynos.length; j++) {
						String[] info = new String[9];
						info[0] = sequence_no;// ????????
						info[1] = assynos[j];// ????????
						info[2] = assyname;// ????????
						info[3] = quantity;// ????
						info[4] = "";// ????
						info[5] = "";// ????
						info[6] = "????????";// ???????? ??????
						if(assynos[j]!=null)
						{
							templist.add(info);
						}						
					}
				}
			}
		}

		return templist;
	}

	/*
	 * ??????????????????????
	 */
	private TCComponentBOMLine getSymmetryState(TCComponentBOMLine linebl, String gwname) throws TCException {
		TCComponentBOMLine ssgwbl = null;
		String ProcLinename = Util.getProperty(linebl, "bl_rev_object_name");
		if (ProcLinename.length() > 1) {
			String rl = ProcLinename.substring(ProcLinename.length() - 2, ProcLinename.length());
			System.out.println("??????????????" + rl);
			if (rl.equals("LH") || rl.equals("RH")) {
				String preLinename = ProcLinename.substring(0, ProcLinename.length() - 2);
				System.out.println("??????????" + ProcLinename);
				ArrayList list = Util.getChildrenByBOMLine(linebl.parent(), "B8_BIWMEProcLineRevision");
				for (int i = 0; i < list.size(); i++) {
					TCComponentBOMLine plinebl = (TCComponentBOMLine) list.get(i);
					String plinename = Util.getProperty(plinebl, "bl_rev_object_name");
					System.out.println("??????????????" + plinename);
					if (!plinename.equals(ProcLinename)) {
						if (plinename.length() > 1
								&& plinename.substring(0, plinename.length() - 2).equals(preLinename)) {
							ArrayList gwlist = Util.getChildrenByBOMLine(plinebl, "B8_BIWMEProcStatRevision");
							for (int j = 0; j < gwlist.size(); j++) {
								TCComponentBOMLine bl = (TCComponentBOMLine) gwlist.get(j);
								String statename = Util.getProperty(bl, "bl_rev_object_name");
								// ????????????????????????????????????????????????????????????????????
								if (gwname.length() > 1) {
									String r2 = gwname.substring(gwname.length() - 2, gwname.length());
									if (r2.equals("LH") || r2.equals("RH")) {
										if (statename.length() > 1) {
											if (statename.substring(0, statename.length() - 2)
													.equals(gwname.substring(0, gwname.length() - 2))) {
												ssgwbl = bl;
												break;
											}
										}
									} else {
										if (statename.equals(gwname)) {
											ssgwbl = bl;
											break;
										}
									}
								} else {
									if (statename.equals(gwname)) {
										ssgwbl = bl;
										break;
									}
								}
							}
						}
					}
				}
			}
		}
		return ssgwbl;
	}

	/*
	 * ????????????????????????????home??,??????00.??????????????????????????00.????????????
	 */
	private void saveFileToFolder(TCComponentItem document, String topfoldername, String childrenFoldername,
			String procName) {
		// TODO Auto-generated method stub
		try {
			TCComponentUser user = session.getUser();
			TCComponentFolder homefolder = user.getHomeFolder();
			TCComponentFolder folder = null;
			TCComponentFolder childrenfolder = null;
			// ????????????????????????????
			AIFComponentContext[] icf = homefolder.getChildren();
			for (AIFComponentContext aif : icf) {
				TCComponent tcc = (TCComponent) aif.getComponent();
				String obejctname = Util.getProperty(tcc, "object_name");
				if (tcc.getType().equals("Folder") && obejctname.equals(topfoldername)) {
					folder = (TCComponentFolder) tcc;
					break;
				}
			}
			if (folder == null) {
				return;
			}
			TCComponentFolderType foldertype = (TCComponentFolderType) session.getTypeComponent("Folder");

			// ????????????????????01.????????????????
			AIFComponentContext[] icf1 = folder.getChildren();
			for (AIFComponentContext aif : icf1) {
				TCComponent tcc = (TCComponent) aif.getComponent();
				String obejctname = Util.getProperty(tcc, "object_name");
				if (tcc.getType().equals("Folder") && obejctname.equals(childrenFoldername)) {
					childrenfolder = (TCComponentFolder) tcc;
					break;
				}
			}
			if (childrenfolder == null) {
				childrenfolder = foldertype.create(childrenFoldername, "", "Folder");
				folder.add("contents", childrenfolder);
				childrenfolder.add("contents", document);
			} else {
				// folder.add("contents", childrenfolder);
				AIFComponentContext[] icf3 = childrenfolder.getChildren();
				// ??????
				if (icf3 != null) {
					for (AIFComponentContext aif : icf3) {
						TCComponent tcc = (TCComponent) aif.getComponent();
						String gwname = Util.getProperty(tcc, "object_name");
						if (gwname.equals(procName)) {
							childrenfolder.remove("contents", tcc);
						}
					}
				}
				childrenfolder.add("contents", document);
			}

		} catch (TCException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
	}

	/*
	 * ??????????????????????????
	 */
	private void getPageNumberManagement(XSSFWorkbook book) {
		// TODO Auto-generated method stub
		int sheetnum = book.getNumberOfSheets();
		for (int i = 0; i < sheetnum; i++) {
			String sheetname = book.getSheetName(i);
			if (sheetname.contains("PSW") || sheetname.contains("RSW")) {
				XSSFSheet sheet = book.getSheetAt(i);

				// ??????12????47??
				int rowStart = 11;
				int rowEnd = 47;
				for (int rowNum = rowStart; rowNum < rowEnd; rowNum++) {
					Row row = (Row) sheet.getRow(rowNum);
					if (null == row) {
						continue;
					}
					Cell cell;
					// ????????
					cell = row.getCell(8);
					String weldno = convertCellValueToString(cell, Cell.CELL_TYPE_STRING);
					if (weldno != null && !weldno.isEmpty()) {
						String[] value = new String[2];
						// ????
						cell = row.getCell(2);
						String pagenumber = convertCellValueToString(cell, Cell.CELL_TYPE_STRING);
						// ??????
						cell = row.getCell(6);
						String dot = convertCellValueToString(cell, Cell.CELL_TYPE_STRING);
						value[0] = pagenumber;
						value[1] = dot;
						notelist.put(weldno, value);
					}

				}
			}
		}
	}

	private static String convertCellValueToString(Cell cell, int type) {
		if (cell == null) {
			return null;
		}
		String returnValue = null;
		switch (type) {
		case Cell.CELL_TYPE_NUMERIC: // ????
			Double doubleValue = cell.getNumericCellValue();
			// ????????????????????????????
			DecimalFormat df = new DecimalFormat("0.00");
			returnValue = df.format(doubleValue);
			break;
		case Cell.CELL_TYPE_STRING: // ??????
			if (cell.getCellType() != Cell.CELL_TYPE_STRING) {
				cell.setCellType(Cell.CELL_TYPE_STRING);
			}
			returnValue = cell.getStringCellValue();
			break;
		case Cell.CELL_TYPE_BOOLEAN: // ????
			Boolean booleanValue = cell.getBooleanCellValue();
			returnValue = booleanValue.toString();
			break;
		case Cell.CELL_TYPE_BLANK: // ????
			break;
		case Cell.CELL_TYPE_FORMULA: // ????
			returnValue = cell.getCellFormula();
			break;
		case Cell.CELL_TYPE_ERROR: // ????
			break;
		default:
			break;
		}
		return returnValue;
	}

	/*
	 * ????RSWSFsheet??????????????
	 */
	private void RSWSFClearSheetContext(XSSFWorkbook book, String name) {
		int sheetnum = 0;
		sheetnum = book.getNumberOfSheets();
		int sheetAtIndex = -1; // RSW????????????
		for (int i = 0; i < sheetnum; i++) {
			String sheetname = book.getSheetName(i);
			if (sheetname.contains(name)) {
				sheetAtIndex = i;
				break;
			}
		}
		if (sheetAtIndex == -1) {
			return;
		}

		// ????????????
		Font font = book.createFont();
		font.setColor((short) 12);// ????????
		font.setFontHeightInPoints((short) 10);
		XSSFCellStyle style = book.createCellStyle();
		style.setBorderBottom(XSSFCellStyle.BORDER_THIN); // ??????
		style.setBorderLeft(XSSFCellStyle.BORDER_THIN);// ??????
		style.setBorderTop(XSSFCellStyle.BORDER_THIN);// ??????
		style.setBorderRight(XSSFCellStyle.BORDER_THIN);// ??????
		style.setVerticalAlignment(XSSFCellStyle.VERTICAL_CENTER);
		style.setAlignment(XSSFCellStyle.ALIGN_CENTER);
		style.setFont(font);

		Font font2 = book.createFont();
		font2.setColor((short) 12);// ????????
		font2.setFontHeightInPoints((short) 18);
		font2.setBoldweight(Font.BOLDWEIGHT_BOLD); // ????????
		XSSFCellStyle style2 = book.createCellStyle();
		style2.setBorderBottom(CellStyle.BORDER_MEDIUM); // ????????
		style2.setBorderLeft(CellStyle.BORDER_MEDIUM); // ????????
		style2.setBorderRight(CellStyle.BORDER_MEDIUM); // ????????
		style2.setBorderTop(CellStyle.BORDER_MEDIUM); // ????????
		style2.setVerticalAlignment(XSSFCellStyle.VERTICAL_CENTER);
		style2.setAlignment(XSSFCellStyle.ALIGN_CENTER);
		style2.setFont(font2);

		// ??????????
		XSSFCellStyle style3 = book.createCellStyle();
		style3.setFillForegroundColor(IndexedColors.PINK.getIndex());
		style3.setFillPattern(CellStyle.SOLID_FOREGROUND);
		style3.setBorderBottom(XSSFCellStyle.BORDER_THIN); // ??????
		style3.setBorderLeft(XSSFCellStyle.BORDER_THIN);// ??????
		style3.setBorderTop(XSSFCellStyle.BORDER_THIN);// ??????
		style3.setBorderRight(XSSFCellStyle.BORDER_THIN);// ??????
		style3.setVerticalAlignment(XSSFCellStyle.VERTICAL_CENTER);
		style3.setAlignment(XSSFCellStyle.ALIGN_CENTER);
		style3.setFont(font);
		// ??????????
		Font font3 = book.createFont();
		font3.setColor((short) 1);// ????????
		font3.setFontHeightInPoints((short) 10);
		XSSFCellStyle style4 = book.createCellStyle();
		style4.setFillForegroundColor(IndexedColors.VIOLET.getIndex());
		style4.setFillPattern(CellStyle.SOLID_FOREGROUND);
		style4.setBorderBottom(XSSFCellStyle.BORDER_THIN); // ??????
		style4.setBorderLeft(XSSFCellStyle.BORDER_THIN);// ??????
		style4.setBorderTop(XSSFCellStyle.BORDER_THIN);// ??????
		style4.setBorderRight(XSSFCellStyle.BORDER_THIN);// ??????
		style4.setVerticalAlignment(XSSFCellStyle.VERTICAL_CENTER);
		style4.setAlignment(XSSFCellStyle.ALIGN_CENTER);
		style4.setFont(font3);
		// ??????????
		Font font4 = book.createFont();
		font4.setColor((short) 1);// ????????
		font4.setFontHeightInPoints((short) 10);
		XSSFCellStyle style5 = book.createCellStyle();
		style5.setFillForegroundColor(IndexedColors.SKY_BLUE.getIndex());
		style5.setFillPattern(CellStyle.SOLID_FOREGROUND);
		style5.setBorderBottom(XSSFCellStyle.BORDER_THIN); // ??????
		style5.setBorderLeft(XSSFCellStyle.BORDER_THIN);// ??????
		style5.setBorderTop(XSSFCellStyle.BORDER_THIN);// ??????
		style5.setBorderRight(XSSFCellStyle.BORDER_THIN);// ??????
		style5.setVerticalAlignment(XSSFCellStyle.VERTICAL_CENTER);
		style5.setAlignment(XSSFCellStyle.ALIGN_CENTER);
		style5.setFont(font4);

		XSSFCellStyle style6 = book.createCellStyle();
		style6.setBorderBottom(XSSFCellStyle.BORDER_THIN); // ??????
		// style6.setBorderLeft(XSSFCellStyle.BORDER_THIN);// ??????
		style6.setBorderTop(XSSFCellStyle.BORDER_THIN);// ??????
		style6.setBorderRight(XSSFCellStyle.BORDER_THIN);// ??????
		style6.setVerticalAlignment(XSSFCellStyle.VERTICAL_CENTER);
		style6.setAlignment(XSSFCellStyle.ALIGN_CENTER);
		style6.setFont(font);

		XSSFCellStyle style8 = book.createCellStyle();
		style8.setBorderBottom(XSSFCellStyle.BORDER_THIN); // ??????
		style8.setBorderLeft(XSSFCellStyle.BORDER_THIN);// ??????
		style8.setBorderTop(XSSFCellStyle.BORDER_THIN);// ??????
		// style8.setBorderRight(XSSFCellStyle.BORDER_THIN);// ??????
		style8.setVerticalAlignment(XSSFCellStyle.VERTICAL_CENTER);
		style8.setAlignment(XSSFCellStyle.ALIGN_CENTER);
		style8.setFont(font);

		// ??????????
		Font font5 = book.createFont();
		font4.setFontHeightInPoints((short) 10);
		XSSFCellStyle style7 = book.createCellStyle();
		style7.setFillForegroundColor(IndexedColors.WHITE.getIndex());
		style7.setFillPattern(CellStyle.SOLID_FOREGROUND);
		style7.setBorderBottom(XSSFCellStyle.BORDER_THIN); // ??????
		style7.setBorderLeft(XSSFCellStyle.BORDER_THIN);// ??????
		style7.setBorderTop(XSSFCellStyle.BORDER_THIN);// ??????
		style7.setBorderRight(XSSFCellStyle.BORDER_THIN);// ??????
		style7.setVerticalAlignment(XSSFCellStyle.VERTICAL_CENTER);
		style7.setAlignment(XSSFCellStyle.ALIGN_CENTER);
		style7.setFont(font5);

		int gcnum = 0;
		for (int i = 0; i < sheetnum; i++) {
			String sheetname = book.getSheetName(i);
			if (sheetname.contains(name)) {
				gcnum++;
			}
		}
		int index = sheetAtIndex + 1;
		// ????sheet????????????????????????????
		index = sheetAtIndex + gcnum;

		// ??????????sheet????????????????????????????????????
		XSSFRow row;
		XSSFCell cell;
		for (int i = sheetAtIndex; i < index; i++) {
			XSSFSheet sheet = book.getSheetAt(i);
			// ????????
			setStringCellAndStyle(sheet, "", 6, 20, style2, Cell.CELL_TYPE_STRING);// ????
			setStringCellAndStyle(sheet, "", 6, 31, style2, Cell.CELL_TYPE_STRING);// ??????
			setStringCellAndStyle(sheet, "", 6, 48, style2, Cell.CELL_TYPE_STRING);// ????????

			for (int j = 0; j < 36; j++) {
				row = sheet.getRow(11 + j);
				if (row != null) {
					cell = row.getCell(8);
					if (cell != null) {
						String weldno = convertCellValueToString(cell, Cell.CELL_TYPE_STRING);
						if (weldno != null && !weldno.isEmpty()) {
							rswsflist.add(weldno);
						}
					}
				}
				setStringCellAndStyle(sheet, "", 11 + j, 2, style, Cell.CELL_TYPE_STRING);
				setStringCellAndStyle(sheet, "", 11 + j, 6, style8, Cell.CELL_TYPE_STRING);
				setStringCellAndStyle(sheet, "", 11 + j, 4, style, Cell.CELL_TYPE_STRING);
				setStringCellAndStyle(sheet, "", 11 + j, 8, style6, Cell.CELL_TYPE_STRING);
				setStringCellAndStyle(sheet, "", 11 + j, 13, style, Cell.CELL_TYPE_STRING);
				setStringCellAndStyle(sheet, "", 11 + j, 16, style, Cell.CELL_TYPE_STRING);
				setStringCellAndStyle(sheet, "", 11 + j, 29, style7, Cell.CELL_TYPE_STRING);
				setStringCellAndStyle(sheet, "", 11 + j, 36, style, 11);
				setStringCellAndStyle(sheet, "", 11 + j, 39, style, Cell.CELL_TYPE_STRING);
				setStringCellAndStyle(sheet, "", 11 + j, 42, style, Cell.CELL_TYPE_STRING);
				setStringCellAndStyle(sheet, "", 11 + j, 55, style7, Cell.CELL_TYPE_STRING);
				setStringCellAndStyle(sheet, "", 11 + j, 62, style, 11);
				setStringCellAndStyle(sheet, "", 11 + j, 65, style, Cell.CELL_TYPE_STRING);
				setStringCellAndStyle(sheet, "", 11 + j, 68, style, Cell.CELL_TYPE_STRING);
				setStringCellAndStyle(sheet, "", 11 + j, 81, style7, Cell.CELL_TYPE_STRING);
				setStringCellAndStyle(sheet, "", 11 + j, 87, style, 11);
				setStringCellAndStyle(sheet, "", 11 + j, 90, style, 10);
				setStringCellAndStyle(sheet, "", 11 + j, 92, style, Cell.CELL_TYPE_STRING);
				setStringCellAndStyle(sheet, "", 11 + j, 94, style, 10);
				setStringCellAndStyle(sheet, "", 11 + j, 96, style, 10);
				setStringCellAndStyle(sheet, "", 11 + j, 98, style, 10);
				setStringCellAndStyle(sheet, "", 11 + j, 100, style, Cell.CELL_TYPE_STRING);
				setStringCellAndStyle(sheet, "", 11 + j, 102, style7, 11);
				setStringCellAndStyle(sheet, "", 11 + j, 105, style, Cell.CELL_TYPE_STRING);
				setStringCellAndStyle(sheet, "", 11 + j, 108, style, Cell.CELL_TYPE_STRING);
				setStringCellAndStyle(sheet, "", 11 + j, 111, style, Cell.CELL_TYPE_STRING);
			}
		}
	}

	/*
	 * ????RSWQDsheet??????????????
	 */
	private void RSWQDClearSheetContext(XSSFWorkbook book, String name) {
		int sheetnum = 0;
		sheetnum = book.getNumberOfSheets();
		int sheetAtIndex = -1; // RSW????????????
		for (int i = 0; i < sheetnum; i++) {
			String sheetname = book.getSheetName(i);
			if (sheetname.contains(name)) {
				sheetAtIndex = i;
				break;
			}
		}
		if (sheetAtIndex == -1) {
			return;
		}

		int gcnum = 0;
		for (int i = 0; i < sheetnum; i++) {
			String sheetname = book.getSheetName(i);
			if (sheetname.contains(name)) {
				gcnum++;
			}
		}
		int index = sheetAtIndex + 1;
		// ????sheet????????????????????????????
		index = sheetAtIndex + gcnum;

		// ????????????
		Font font = book.createFont();
		font.setColor((short) 12);// ????????
		font.setFontHeightInPoints((short) 10);
		XSSFCellStyle style = book.createCellStyle();
		style.setBorderBottom(XSSFCellStyle.BORDER_THIN); // ??????
		style.setBorderLeft(XSSFCellStyle.BORDER_THIN);// ??????
		style.setBorderTop(XSSFCellStyle.BORDER_THIN);// ??????
		style.setBorderRight(XSSFCellStyle.BORDER_THIN);// ??????
		style.setVerticalAlignment(XSSFCellStyle.VERTICAL_CENTER);
		style.setAlignment(XSSFCellStyle.ALIGN_CENTER);
		style.setFont(font);

		Font font2 = book.createFont();
		font2.setColor((short) 12);// ????????
		font2.setFontHeightInPoints((short) 18);
		font2.setBoldweight(Font.BOLDWEIGHT_BOLD); // ????????
		XSSFCellStyle style2 = book.createCellStyle();
		style2.setBorderBottom(CellStyle.BORDER_MEDIUM); // ????????
		style2.setBorderLeft(CellStyle.BORDER_MEDIUM); // ????????
		style2.setBorderRight(CellStyle.BORDER_MEDIUM); // ????????
		style2.setBorderTop(CellStyle.BORDER_MEDIUM); // ????????
		style2.setVerticalAlignment(XSSFCellStyle.VERTICAL_CENTER);
		style2.setAlignment(XSSFCellStyle.ALIGN_CENTER);
		style2.setFont(font2);

		XSSFCellStyle style3 = book.createCellStyle();
		style3.setBorderBottom(XSSFCellStyle.BORDER_THIN); // ??????
		// style3.setBorderLeft(XSSFCellStyle.BORDER_THIN);// ??????
		style3.setBorderTop(XSSFCellStyle.BORDER_THIN);// ??????
		style3.setBorderRight(XSSFCellStyle.BORDER_THIN);// ??????
		style3.setVerticalAlignment(XSSFCellStyle.VERTICAL_CENTER);
		style3.setAlignment(XSSFCellStyle.ALIGN_CENTER);
		style3.setFont(font);

		XSSFCellStyle style4 = book.createCellStyle();
		style4.setBorderBottom(XSSFCellStyle.BORDER_THIN); // ??????
		style4.setBorderLeft(XSSFCellStyle.BORDER_THIN);// ??????
		style4.setBorderTop(XSSFCellStyle.BORDER_THIN);// ??????
		// style4.setBorderRight(XSSFCellStyle.BORDER_THIN);// ??????
		style4.setVerticalAlignment(XSSFCellStyle.VERTICAL_CENTER);
		style4.setAlignment(XSSFCellStyle.ALIGN_CENTER);
		style4.setFont(font);

		// ??????????sheet????????????????????????????????????
		XSSFRow row;
		XSSFCell cell;
		for (int i = sheetAtIndex; i < index; i++) {
			XSSFSheet sheet = book.getSheetAt(i);

			// ????????
			setStringCellAndStyle(sheet, "", 6, 19, style2, Cell.CELL_TYPE_STRING);// ????
			setStringCellAndStyle(sheet, "", 6, 30, style2, Cell.CELL_TYPE_STRING);// ??????
			setStringCellAndStyle(sheet, "", 6, 47, style2, Cell.CELL_TYPE_STRING);// ????????

			for (int j = 0; j < 36; j++) {
				row = sheet.getRow(11 + j);
				if (row != null) {
					cell = row.getCell(8);
					if (cell != null) {
						String weldno = convertCellValueToString(cell, Cell.CELL_TYPE_STRING);
						if (weldno != null && !weldno.isEmpty()) {
							rswqdlist.add(weldno);
						}
					}
				}
				setStringCellAndStyle(sheet, "", 11 + j, 2, style, Cell.CELL_TYPE_STRING);
				setStringCellAndStyle(sheet, "", 11 + j, 6, style4, Cell.CELL_TYPE_STRING);
				setStringCellAndStyle(sheet, "", 11 + j, 4, style, Cell.CELL_TYPE_STRING);
				setStringCellAndStyle(sheet, "", 11 + j, 8, style3, Cell.CELL_TYPE_STRING);
				setStringCellAndStyle(sheet, "", 11 + j, 13, style, Cell.CELL_TYPE_STRING);
				setStringCellAndStyle(sheet, "", 11 + j, 16, style, Cell.CELL_TYPE_STRING);
				setStringCellAndStyle(sheet, "", 11 + j, 29, style, Cell.CELL_TYPE_STRING);
				setStringCellAndStyle(sheet, "", 11 + j, 36, style, 11);
				setStringCellAndStyle(sheet, "", 11 + j, 39, style, Cell.CELL_TYPE_STRING);
				setStringCellAndStyle(sheet, "", 11 + j, 42, style, Cell.CELL_TYPE_STRING);
				setStringCellAndStyle(sheet, "", 11 + j, 55, style, Cell.CELL_TYPE_STRING);
				setStringCellAndStyle(sheet, "", 11 + j, 62, style, 11);
				setStringCellAndStyle(sheet, "", 11 + j, 65, style, Cell.CELL_TYPE_STRING);
				setStringCellAndStyle(sheet, "", 11 + j, 68, style, Cell.CELL_TYPE_STRING);
				setStringCellAndStyle(sheet, "", 11 + j, 81, style, Cell.CELL_TYPE_STRING);
				setStringCellAndStyle(sheet, "", 11 + j, 88, style, 11);
				setStringCellAndStyle(sheet, "", 11 + j, 91, style, 10);
				setStringCellAndStyle(sheet, "", 11 + j, 93, style, Cell.CELL_TYPE_STRING);
				setStringCellAndStyle(sheet, "", 11 + j, 95, style, 10);
				setStringCellAndStyle(sheet, "", 11 + j, 97, style, 10);
				setStringCellAndStyle(sheet, "", 11 + j, 99, style, 10);
				setStringCellAndStyle(sheet, "", 11 + j, 102, style, 11);
				setStringCellAndStyle(sheet, "", 11 + j, 105, style, 10);
				setStringCellAndStyle(sheet, "", 11 + j, 108, style, 10);
				setStringCellAndStyle(sheet, "", 11 + j, 111, style, 10);
			}
		}
	}

	/*
	 * ????PSWsheet??????????????
	 */
	private void PSWClearSheetContext(XSSFWorkbook book, String name) {
		int sheetnum = 0;
		sheetnum = book.getNumberOfSheets();
		int sheetAtIndex = -1; // PSW????????
		for (int i = 0; i < sheetnum; i++) {
			String sheetname = book.getSheetName(i);
			if (sheetname.contains(name) && !sheetname.contains("????")) {
				sheetAtIndex = i;
				break;
			}
		}
		if (sheetAtIndex == -1) {
			return;
		}

		int gcnum = 0;
		for (int i = 0; i < sheetnum; i++) {
			String sheetname = book.getSheetName(i);
			if (sheetname.contains(name) && !sheetname.contains("????")) {
				gcnum++;
			}
		}
		int index = sheetAtIndex + 1;
		// ????sheet????????????????????????????
		index = sheetAtIndex + gcnum;

		// ????????????
		Font font = book.createFont();
		font.setColor((short) 12);// ????????
		font.setFontHeightInPoints((short) 10);
		XSSFCellStyle style = book.createCellStyle();
		style.setBorderBottom(XSSFCellStyle.BORDER_THIN); // ??????
		style.setBorderLeft(XSSFCellStyle.BORDER_THIN);// ??????
		style.setBorderTop(XSSFCellStyle.BORDER_THIN);// ??????
		style.setBorderRight(XSSFCellStyle.BORDER_THIN);// ??????
		style.setVerticalAlignment(XSSFCellStyle.VERTICAL_CENTER);
		style.setAlignment(XSSFCellStyle.ALIGN_CENTER);
		style.setFont(font);

		Font font2 = book.createFont();
		font2.setColor((short) 12);// ????????
		font2.setFontHeightInPoints((short) 11);
		font2.setBoldweight(Font.BOLDWEIGHT_BOLD); // ????????
		XSSFCellStyle style2 = book.createCellStyle();
		style2.setBorderBottom(CellStyle.BORDER_MEDIUM); // ????????
		style2.setBorderLeft(CellStyle.BORDER_MEDIUM); // ????????
		style2.setBorderRight(CellStyle.BORDER_MEDIUM); // ????????
		style2.setBorderTop(CellStyle.BORDER_MEDIUM); // ????????
		style2.setVerticalAlignment(XSSFCellStyle.VERTICAL_CENTER);
		style2.setAlignment(XSSFCellStyle.ALIGN_CENTER);
		style2.setFont(font2);

		XSSFCellStyle style22 = book.createCellStyle();
		style22.setBorderBottom(CellStyle.BORDER_MEDIUM); // ????????
		style22.setBorderLeft(CellStyle.BORDER_THIN); // ????????
		style22.setBorderRight(CellStyle.BORDER_THIN); // ????????
		style22.setBorderTop(CellStyle.BORDER_THIN); //
		style22.setVerticalAlignment(XSSFCellStyle.VERTICAL_CENTER);
		style22.setAlignment(XSSFCellStyle.ALIGN_CENTER);
		style22.setFont(font2);

		XSSFCellStyle style20 = book.createCellStyle();
		style20.setBorderBottom(CellStyle.BORDER_MEDIUM); // ????????
		style20.setBorderLeft(CellStyle.BORDER_MEDIUM); // ????????
		style20.setBorderRight(CellStyle.BORDER_THIN); // ????????
		style20.setBorderTop(CellStyle.BORDER_THIN); //
		style20.setVerticalAlignment(XSSFCellStyle.VERTICAL_CENTER);
		style20.setAlignment(XSSFCellStyle.ALIGN_CENTER);
		style20.setFont(font2);

		XSSFCellStyle style21 = book.createCellStyle();
		style21.setBorderBottom(CellStyle.BORDER_MEDIUM); // ????????
		style21.setBorderLeft(CellStyle.BORDER_THIN); // ????????
		style21.setBorderRight(CellStyle.BORDER_MEDIUM); // ????????
		style21.setBorderTop(CellStyle.BORDER_THIN); //
		style21.setVerticalAlignment(XSSFCellStyle.VERTICAL_CENTER);
		style21.setAlignment(XSSFCellStyle.ALIGN_CENTER);
		style21.setFont(font2);

		Font font3 = book.createFont();
		font3.setColor((short) 12);// ????????
		font3.setFontHeightInPoints((short) 18);
		font3.setBoldweight(Font.BOLDWEIGHT_BOLD); // ????????
		XSSFCellStyle style3 = book.createCellStyle();
		style3.setBorderBottom(CellStyle.BORDER_MEDIUM); // ????????
		style3.setBorderLeft(CellStyle.BORDER_MEDIUM); // ????????
		style3.setBorderRight(CellStyle.BORDER_MEDIUM); // ????????
		style3.setBorderTop(CellStyle.BORDER_MEDIUM); // ????????
		style3.setVerticalAlignment(XSSFCellStyle.VERTICAL_CENTER);
		style3.setAlignment(XSSFCellStyle.ALIGN_CENTER);
		style3.setFont(font3);

		XSSFCellStyle style4 = book.createCellStyle();
		style4.setBorderBottom(XSSFCellStyle.BORDER_THIN); // ??????
		// style4.setBorderLeft(XSSFCellStyle.BORDER_THIN);// ??????
		style4.setBorderTop(XSSFCellStyle.BORDER_THIN);// ??????
		style4.setBorderRight(XSSFCellStyle.BORDER_THIN);// ??????
		style4.setVerticalAlignment(XSSFCellStyle.VERTICAL_CENTER);
		style4.setAlignment(XSSFCellStyle.ALIGN_CENTER);
		style4.setFont(font);

		XSSFCellStyle style5 = book.createCellStyle();
		style5.setBorderBottom(XSSFCellStyle.BORDER_THIN); // ??????
		style5.setBorderLeft(XSSFCellStyle.BORDER_THIN);// ??????
		style5.setBorderTop(XSSFCellStyle.BORDER_THIN);// ??????
		// style5.setBorderRight(XSSFCellStyle.BORDER_THIN);// ??????
		style5.setVerticalAlignment(XSSFCellStyle.VERTICAL_CENTER);
		style5.setAlignment(XSSFCellStyle.ALIGN_CENTER);
		style5.setFont(font);

		// ??????????sheet????????????????????????????????????
		XSSFRow row;
		XSSFCell cell;
		for (int i = sheetAtIndex; i < index; i++) {
			XSSFSheet sheet = book.getSheetAt(i);
			// ????????

			setStringCellAndStyle(sheet, "", 5, 8, style2, Cell.CELL_TYPE_STRING);// ??????????
			setStringCellAndStyle(sheet, "", 5, 19, style3, Cell.CELL_TYPE_STRING);// ????????
			setStringCellAndStyle(sheet, "", 7, 36, style20, Cell.CELL_TYPE_STRING);// ??????
			setStringCellAndStyle(sheet, "", 7, 42, style22, Cell.CELL_TYPE_STRING);// ????????
			setStringCellAndStyle(sheet, "", 7, 48, style22, Cell.CELL_TYPE_STRING);// ????????
			setStringCellAndStyle(sheet, "", 7, 54, style22, Cell.CELL_TYPE_STRING);// ???? ????????
			setStringCellAndStyle(sheet, "", 7, 60, style22, Cell.CELL_TYPE_STRING);// ???? ????????
			setStringCellAndStyle(sheet, "", 7, 66, style22, Cell.CELL_TYPE_STRING);// ??????????
			setStringCellAndStyle(sheet, "", 7, 72, style22, Cell.CELL_TYPE_STRING);// ???? ????????
			setStringCellAndStyle(sheet, "", 7, 78, style22, Cell.CELL_TYPE_STRING);// ???? ????????
			setStringCellAndStyle(sheet, "", 7, 84, style22, Cell.CELL_TYPE_STRING);// ??????????
			setStringCellAndStyle(sheet, "", 7, 90, style22, Cell.CELL_TYPE_STRING);// ???? ????????
			setStringCellAndStyle(sheet, "", 7, 96, style22, Cell.CELL_TYPE_STRING);// ???? ????????
			setStringCellAndStyle(sheet, "", 7, 102, style21, Cell.CELL_TYPE_STRING);// ????

			for (int j = 0; j < 36; j++) {
				row = sheet.getRow(11 + j);
				if (row != null) {
					cell = row.getCell(8);
					if (cell != null) {
						String weldno = convertCellValueToString(cell, Cell.CELL_TYPE_STRING);
						if (weldno != null && !weldno.isEmpty()) {
							pswlist.add(weldno);
						}
					}
				}
				setStringCellAndStyle(sheet, "", 11 + j, 2, style, Cell.CELL_TYPE_STRING);
				setStringCellAndStyle(sheet, "", 11 + j, 6, style5, Cell.CELL_TYPE_STRING);
				setStringCellAndStyle(sheet, "", 11 + j, 4, style, Cell.CELL_TYPE_STRING);
				setStringCellAndStyle(sheet, "", 11 + j, 8, style4, Cell.CELL_TYPE_STRING);
				setStringCellAndStyle(sheet, "", 11 + j, 13, style, Cell.CELL_TYPE_STRING);
				setStringCellAndStyle(sheet, "", 11 + j, 16, style, Cell.CELL_TYPE_STRING);
				setStringCellAndStyle(sheet, "", 11 + j, 29, style, Cell.CELL_TYPE_STRING);
				setStringCellAndStyle(sheet, "", 11 + j, 36, style, 11);
				setStringCellAndStyle(sheet, "", 11 + j, 39, style, Cell.CELL_TYPE_STRING);
				setStringCellAndStyle(sheet, "", 11 + j, 42, style, Cell.CELL_TYPE_STRING);
				setStringCellAndStyle(sheet, "", 11 + j, 55, style, Cell.CELL_TYPE_STRING);
				setStringCellAndStyle(sheet, "", 11 + j, 62, style, 11);
				setStringCellAndStyle(sheet, "", 11 + j, 65, style, Cell.CELL_TYPE_STRING);
				setStringCellAndStyle(sheet, "", 11 + j, 68, style, Cell.CELL_TYPE_STRING);
				setStringCellAndStyle(sheet, "", 11 + j, 81, style, Cell.CELL_TYPE_STRING);
				setStringCellAndStyle(sheet, "", 11 + j, 88, style, 11);
				setStringCellAndStyle(sheet, "", 11 + j, 91, style, 10);
				setStringCellAndStyle(sheet, "", 11 + j, 93, style, Cell.CELL_TYPE_STRING);
				setStringCellAndStyle(sheet, "", 11 + j, 95, style, 10);
				setStringCellAndStyle(sheet, "", 11 + j, 97, style, 10);
				setStringCellAndStyle(sheet, "", 11 + j, 99, style, 10);
				setStringCellAndStyle(sheet, "", 11 + j, 102, style, 11);
				setStringCellAndStyle(sheet, "", 11 + j, 105, style, 10);
				setStringCellAndStyle(sheet, "", 11 + j, 108, style, 10);
				setStringCellAndStyle(sheet, "", 11 + j, 111, style, 10);
			}
		}
	}

	/*
	 * ????????????????sheet??
	 */
	private String CopySheet(XSSFWorkbook book, String name, int num) {
		String shname = "";
//		if (!updateflag) 
//		{
//			for(int i=0;i<deletelist.size();i++)
//			{
//				if(deletelist.get(i).toString().contains(name) && !deletelist.get(i).toString().contains("????"))
//				{
//					return null;
//				}
//			}			
//		}	
		int sheetnum = 0;
		sheetnum = book.getNumberOfSheets();
		int sheetAtIndex = -1; // PSW????????
		for (int i = 0; i < sheetnum; i++) {
			String sheetname = book.getSheetName(i);
			if (sheetname.contains(name) && !sheetname.contains("????")) {
				sheetAtIndex = i;
				break;
			}
		}
		if (sheetAtIndex == -1) {
			return null;
		}
		if (updateflag) {
			XSSFSheet sh = book.getSheetAt(sheetAtIndex + num);
			if (sh == null) {
				XSSFSheet sheet = book.cloneSheet(sheetAtIndex);
				shname = sheet.getSheetName();
				book.setSheetOrder(sheet.getSheetName(), sheetAtIndex + 1);
			} else {
				if (sh.getSheetName().contains(name) && !sh.getSheetName().contains("????")) {
					shname = sh.getSheetName();
				} else {
					XSSFSheet sheet = book.cloneSheet(sheetAtIndex);
					shname = sheet.getSheetName();
					book.setSheetOrder(sheet.getSheetName(), sheetAtIndex + 1);
				}
			}
		} else {
			if (num == 0) {
				shname = book.getSheetName(sheetAtIndex);
			} else {
				XSSFSheet sheet = book.cloneSheet(sheetAtIndex);
				shname = sheet.getSheetName();
				book.setSheetOrder(sheet.getSheetName(), sheetAtIndex + num);
			}
		}

		return shname;
	}

	/*
	 * ????????????????
	 */
	private void ProcessingInstallationIcon(XSSFWorkbook book, TCComponentBOMLine gwbl) throws TCException {
		// TODO Auto-generated method stub
		// ??????????????????sheet
		int sheetnum = 0;
		sheetnum = book.getNumberOfSheets();
		int sheetAtIndex1 = -1; // ????????????????
		int sheetAtIndex2 = -1; // ????????????
		for (int i = 0; i < sheetnum; i++) {
			String sheetname = book.getSheetName(i);
			if (sheetname.contains("????????")) {
				sheetAtIndex1 = i;
			}
			if (sheetname.contains("????")) {
				sheetAtIndex2 = i;
			}
		}
		if (sheetAtIndex1 == -1 && sheetAtIndex2 == -1) {
			return;
		}
		/**************************************************/
		// ????????????????????????????????????????????????
		if (updateflag) {
//			XSSFSheet sheet = book.getSheetAt(sheetAtIndex1);
//			List<String> delPicturesList = ReportUtils.removePictrues07((XSSFSheet) sheet, (XSSFWorkbook) book, 3, 10,
//					100, 115);
//			System.out.println("-----------????????????????-----------");
//			for (String name : delPicturesList) {
//				System.out.println(name);
//			}
//
//			XSSFSheet sheet2 = book.getSheetAt(sheetAtIndex2);
//			List<String> delPicturesList2 = ReportUtils.removePictrues07((XSSFSheet) sheet2, (XSSFWorkbook) book, 3, 10,
//					100, 115);
//			System.out.println("-----------????????????????-----------");
//			for (String name : delPicturesList2) {
//				System.out.println(name);
//			}

		}
		/**************************************************/

		// ????????????????
		ArrayList tjlist = Util.getChildrenByBOMLine(gwbl, "B8_BIWOperationRevision");
		if (tjlist != null && !tjlist.isEmpty()) {
			String b8_TorqueImptLevel = "";
			for (int i = 0; i < tjlist.size(); i++) {
				TCComponentBOMLine bl = (TCComponentBOMLine) tjlist.get(i);
				TCComponentItemRevision rev = bl.getItemRevision();
				b8_TorqueImptLevel = rev.getProperty("b8_TorqueImptLevel");
				System.out.println(rev.getProperty("b8_TorqueImptLevel"));
				if (b8_TorqueImptLevel != null && !b8_TorqueImptLevel.isEmpty()) {
					break;
				}
			}
			if (b8_TorqueImptLevel != null && !b8_TorqueImptLevel.isEmpty()) {
				if (sheetAtIndex1 != -1) {
					XSSFSheet sheet = book.getSheetAt(sheetAtIndex1);
					InputStream is = null;
					if (b8_TorqueImptLevel.trim().equals("A")) {
						is = this.getClass().getResourceAsStream("/com/dfl/report/imags/A.png");
					}
					if (b8_TorqueImptLevel.trim().equals("B")) {
						is = this.getClass().getResourceAsStream("/com/dfl/report/imags/B.png");
					}
					if (is != null) {
						writepicturetosheet(book, sheet, is, 105, 5, 111, 9);
					}
				}
				if (sheetAtIndex2 != -1) {
					XSSFSheet sheet = book.getSheetAt(sheetAtIndex2);
					InputStream is = null;
					if (b8_TorqueImptLevel.trim().equals("A")) {
						is = this.getClass().getResourceAsStream("/com/dfl/report/imags/A.png");
					}
					if (b8_TorqueImptLevel.trim().equals("B")) {
						is = this.getClass().getResourceAsStream("/com/dfl/report/imags/B.png");
					}
					if (is != null) {
						writepicturetosheet(book, sheet, is, 105, 5, 111, 9);
					}
				}
			}

		}
	}

	/*
	 * ????????????????
	 */
	private void ProcessingGlueIcon(XSSFWorkbook book, TCComponentBOMLine gwbl) throws TCException {
		// TODO Auto-generated method stub
		// ????????sheet
		int sheetnum = 0;
		sheetnum = book.getNumberOfSheets();
		int sheetAtIndex = -1; // ????????????
		for (int i = 0; i < sheetnum; i++) {
			String sheetname = book.getSheetName(i);
			if (sheetname.contains("????")) {
				sheetAtIndex = i;
				break;
			}
		}
		if (sheetAtIndex == -1) {
			return;
		}

		/**************************************************/
		// ????????????????????????????????????????????????
		if (updateflag) {
//			XSSFSheet sheet = book.getSheetAt(sheetAtIndex);
//			List<String> delPicturesList = ReportUtils.removePictrues07((XSSFSheet) sheet, (XSSFWorkbook) book, 3, 10,
//					100, 115);
//			System.out.println("-----------????????????????-----------");
//			for (String name : delPicturesList) {
//				System.out.println(name);
//			}
		}
		/**************************************************/

		// ????????????????
		ArrayList tjlist = Util.getChildrenByBOMLine(gwbl, "B8_BIWArcWeldOPRevision");// B8_BIWPaintOPRevision
		if (tjlist != null && !tjlist.isEmpty()) {
			String b8_GlueFeature = "";
			for (int i = 0; i < tjlist.size(); i++) {
				TCComponentBOMLine bl = (TCComponentBOMLine) tjlist.get(i);
				TCComponentItemRevision rev = bl.getItemRevision();
				b8_GlueFeature = rev.getProperty("b8_GlueFeature");
				System.out.println(rev.getProperty("b8_GlueFeature"));
				if (b8_GlueFeature != null && !b8_GlueFeature.isEmpty()) {
					break;
				}
			}
			if (b8_GlueFeature != null && !b8_GlueFeature.isEmpty()) {
				XSSFSheet sheet = book.getSheetAt(sheetAtIndex);
				InputStream is = null;
				if (b8_GlueFeature.trim().equals("????")) {
					is = this.getClass().getResourceAsStream("/com/dfl/report/imags/SM.png");
				}
				if (b8_GlueFeature.trim().equals("????")) {
					is = this.getClass().getResourceAsStream("/com/dfl/report/imags/FX.png");
				}
				if (is != null) {
					writepicturetosheet(book, sheet, is, 105, 5, 111, 9);
				}
			}

		}
	}

	/*
	 * ??????????????????
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
	 * ????????????????
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

	/*
	 * ??????????????????????????????????
	 */
	private List<WeldPointBoardInformation> getBoardInformation(List<WeldPointBoardInformation> baseinfolist,
			ArrayList hdlist) {
		// TODO Auto-generated method stub

		List<WeldPointBoardInformation> totalinfo = new ArrayList<WeldPointBoardInformation>();
		if (baseinfolist != null) {
			for (int i = 0; i < hdlist.size(); i++) {
				TCComponentBOMLine hdbl = (TCComponentBOMLine) hdlist.get(i);
				String weldno = Util.getProperty(hdbl, "bl_rev_object_name");
				for (int j = 0; j < baseinfolist.size(); j++) {
					WeldPointBoardInformation wb = baseinfolist.get(j);
					if (wb.getWeldno() != null && weldno != null && wb.getWeldno().equals(weldno)) {
						totalinfo.add(wb);
						break; // ??????????????????????????????????????
					}
				}
			}
		} else {
			System.out.println("??????????????????");
		}

		return totalinfo;
	}

	/*
	 * PSW????????
	 */
	private String PSWinformationProcessing(XSSFWorkbook book, TCComponentBOMLine bl, String name,
			Map<String, TCComponentBOMLine> symmetrymap) throws TCException {
		// TODO Auto-generated method stub
				
		String error = "";
		// ????PSWsheet
		int sheetnum = 0;
		sheetnum = book.getNumberOfSheets();
		int sheetAtIndex = -1; // PSW????????
		for (int i = 0; i < sheetnum; i++) {
			String sheetname = book.getSheetName(i);
			if (sheetname.contains(name)) {
				sheetAtIndex = i;
				break;
			}
		}
		if (sheetAtIndex == -1) {
			return error;
		}
		// ??????????????????????
		ArrayList gunlist = new ArrayList();
		ArrayList hdlist = new ArrayList();

		String guntypename = Util.getObjectDisplayName(session, "B8_BIWGun");
		String[] propertys = new String[] { "bl_item_object_type", "bl_item_object_type" };
		String[] values = new String[] { guntypename, guntypename };

		String weldtypename = Util.getObjectDisplayName(session, "WeldPoint");
		String[] propertys2 = new String[] { "bl_item_object_type", "bl_item_object_type" };
		String[] values2 = new String[] { weldtypename, weldtypename };

		gunlist = Util.searchBOMLine(bl, "OR", propertys, "==", values);
		hdlist = Util.searchBOMLine(bl, "OR", propertys2, "==", values2);

		// ??????????????SOP????????????SOP??????????????????
		boolean sopflag = getIsSOPAfter();
		String Discretename = Util.getProperty(bl, "bl_rev_object_name");

		// ????????????????????????
		String TransformerNumber = "";// ??????????
		String Guncode = "";// ????????
		String ElectrodeVol = "";// ???????? bl_B8_BIWGunRevision_b8_ElectrodeVol
		TCComponentItemRevision blrev = null;
		TransformerNumber = Util.getProperty(bl.getItemRevision(), "b8_AdapterModel");
		if (gunlist.size() > 0) {
			TCComponentBOMLine gunbl = (TCComponentBOMLine) gunlist.get(0);
			blrev = gunbl.getItemRevision();
//			TransformerNumber = Util.getProperty(blrev, "b8_AdapterModel");
//			Guncode = Util.getProperty(blrev, "b8_Model");
			ElectrodeVol = Util.getProperty(blrev, "b8_ElectrodeVol");
		}
		if (sopflag) {
			if (TransformerNumber == null || TransformerNumber.isEmpty()) {
				error = "??????????????????????????????????????";
				return error;
			}
		}
		// ??????????????????????????????????????????????
		String[] nameArr = Discretename.split("\\\\");
		TransformerNumber = nameArr[0];
		if (nameArr.length > 1) {
			Guncode = nameArr[1];
		}
		if (sopflag) {
			if ("TR NO".equals(TransformerNumber)) {
				error = "??????????????????????????????????????";
				return error;
			}
		}
		// ??????????????????????????????????
		List<WeldPointBoardInformation> hdinfo = new ArrayList<WeldPointBoardInformation>();// ????????????

		hdinfo = getBoardInformation(baseinfolist, hdlist);

		Map<String, String[]> hjmap = getHanJieParater(hdlist);

		if (hdinfo == null || hdinfo.size() < 1) {
			System.out.println("????????????????????????????");
			return error;
		}

		System.out.println("hdinfo: " + hdinfo.size());

		// ????????????????????
		Comparator comparator = getComParatorByfirstpart();
		Collections.sort(hdinfo, comparator);
		// ??????????????????????
		Comparator comparator1 = getComParatorBySecondpart();
		Collections.sort(hdinfo, comparator1);

		// ????????????????????
		Comparator comparator11 = getComParatorByThistypart();
		Collections.sort(hdinfo, comparator11);

		// ??????????????????????3????????????2????
		Comparator comparator2 = getComParatorBylayersnum();
		Collections.sort(hdinfo, comparator2);

		// ??????????????????????????????????????????????????????????????????????????????????sheet??????

		// ??????????????????????????????????
		List<WeldPointBoardInformation> symmetryhdinfo = new ArrayList<WeldPointBoardInformation>();// ????????????
		ArrayList symmetryhdlist = new ArrayList();
		if (symmetrymap.containsKey(Discretename)) {
			TCComponentBOMLine symmetrybl = symmetrymap.get(Discretename);
			symmetryhdlist = Util.searchBOMLine(symmetrybl, "OR", propertys2, "==", values2);
			symmetryhdinfo = getBoardInformation(baseinfolist, symmetryhdlist);

			// ????????????????????
			Comparator comparators1 = getComParatorByfirstpart();
			Collections.sort(symmetryhdinfo, comparators1);
			// ??????????????????????
			Comparator comparators2 = getComParatorBySecondpart();
			Collections.sort(symmetryhdinfo, comparators2);

			// ????????????????????
			Comparator comparators3 = getComParatorByThistypart();
			Collections.sort(symmetryhdinfo, comparators3);

			// ??????????????????????3????????????2????
			Comparator comparators4 = getComParatorBylayersnum();
			Collections.sort(symmetryhdinfo, comparators4);

			// ??????????????????
			getSymmetryHanJieParater(symmetryhdlist, hjmap);
		}

		// ????????????????????????,??36????????????
		int hdsum = hdlist.size() + symmetryhdlist.size();
		int page = hdsum / 36 + 1;

		// ??????????????????????sheet????????????????
		if (hdsum % 36 == 0) {
			if (page > 1) {
				page = page - 1;
			}
		}

		// ????page????1????????????sheet??
		int index = sheetAtIndex + 1;

		// ????????????
		Font font = book.createFont();
		font.setColor((short) 12);// ????????
		font.setFontHeightInPoints((short) 9);
		XSSFCellStyle style = book.createCellStyle();
		style.setBorderBottom(XSSFCellStyle.BORDER_THIN); // ??????
		style.setBorderLeft(XSSFCellStyle.BORDER_THIN);// ??????
		style.setBorderTop(XSSFCellStyle.BORDER_THIN);// ??????
		style.setBorderRight(XSSFCellStyle.BORDER_THIN);// ??????
		style.setVerticalAlignment(XSSFCellStyle.VERTICAL_CENTER);
		style.setAlignment(XSSFCellStyle.ALIGN_CENTER);
		style.setFont(font);

		Font font2 = book.createFont();
		font2.setColor((short) 12);// ????????
		font2.setFontName("MS PGothic");
		font2.setFontHeightInPoints((short) 11);
		font2.setBoldweight(Font.BOLDWEIGHT_BOLD); // ????????
		XSSFCellStyle style2 = book.createCellStyle();
		style2.setBorderBottom(CellStyle.BORDER_MEDIUM); // ????????
		style2.setBorderLeft(CellStyle.BORDER_MEDIUM); // ????????
		style2.setBorderRight(CellStyle.BORDER_MEDIUM); // ????????
		style2.setBorderTop(CellStyle.BORDER_MEDIUM); // ????????
		style2.setVerticalAlignment(XSSFCellStyle.VERTICAL_CENTER);
		style2.setAlignment(XSSFCellStyle.ALIGN_CENTER);
		style2.setFont(font2);

		XSSFCellStyle style22 = book.createCellStyle();
//		style22.setBorderBottom(CellStyle.BORDER_MEDIUM); // ????????
//		style22.setBorderLeft(CellStyle.BORDER_THIN); // ????????
//		style22.setBorderRight(CellStyle.BORDER_THIN); // ????????
//		style22.setBorderTop(CellStyle.BORDER_THIN); //
		style22.setVerticalAlignment(XSSFCellStyle.VERTICAL_CENTER);
		style22.setAlignment(XSSFCellStyle.ALIGN_CENTER);
		style22.setFont(font2);

		XSSFCellStyle style20 = book.createCellStyle();
//		style20.setBorderBottom(CellStyle.BORDER_MEDIUM); // ????????
//		style20.setBorderLeft(CellStyle.BORDER_MEDIUM); // ????????
//		style20.setBorderRight(CellStyle.BORDER_THIN); // ????????
//		style20.setBorderTop(CellStyle.BORDER_THIN); //
		style20.setVerticalAlignment(XSSFCellStyle.VERTICAL_CENTER);
		style20.setAlignment(XSSFCellStyle.ALIGN_CENTER);
		style20.setFont(font2);

		XSSFCellStyle style21 = book.createCellStyle();
//		style21.setBorderBottom(CellStyle.BORDER_MEDIUM); // ????????
//		style21.setBorderLeft(CellStyle.BORDER_THIN); // ????????
//		style21.setBorderRight(CellStyle.BORDER_MEDIUM); // ????????
//		style21.setBorderTop(CellStyle.BORDER_THIN); //
		style21.setVerticalAlignment(XSSFCellStyle.VERTICAL_CENTER);
		style21.setAlignment(XSSFCellStyle.ALIGN_CENTER);
		style21.setFont(font2);

		Font font222 = book.createFont();
		// font2.setColor((short) 10);// ????????
		font222.setFontName("MS PGothic");
		font222.setFontHeightInPoints((short) 10);

		XSSFCellStyle style202 = book.createCellStyle();
//		style202.setBorderBottom(CellStyle.BORDER_THIN); // ????????
//		style202.setBorderLeft(CellStyle.BORDER_MEDIUM); // ????????
//		style202.setBorderRight(CellStyle.BORDER_THIN); // ????????
//		style202.setBorderTop(CellStyle.BORDER_MEDIUM); //
		style202.setVerticalAlignment(XSSFCellStyle.VERTICAL_CENTER);
		style202.setAlignment(XSSFCellStyle.ALIGN_CENTER);
		style202.setWrapText(true);
		style202.setFont(font222);

		XSSFCellStyle style212 = book.createCellStyle();
//		style212.setBorderBottom(CellStyle.BORDER_THIN); // ????????
//		style212.setBorderLeft(CellStyle.BORDER_THIN); // ????????
//		style212.setBorderRight(CellStyle.BORDER_MEDIUM); // ????????
//		style212.setBorderTop(CellStyle.BORDER_MEDIUM); //
		style212.setVerticalAlignment(XSSFCellStyle.VERTICAL_CENTER);
		style212.setAlignment(XSSFCellStyle.ALIGN_CENTER);
		style212.setWrapText(true);
		style212.setFont(font222);

		XSSFCellStyle style222 = book.createCellStyle();
//		style222.setBorderBottom(CellStyle.BORDER_THIN); // ????????
//		style222.setBorderLeft(CellStyle.BORDER_THIN); // ????????
//		style222.setBorderRight(CellStyle.BORDER_THIN); // ????????
//		style222.setBorderTop(CellStyle.BORDER_MEDIUM); //
		style222.setVerticalAlignment(XSSFCellStyle.VERTICAL_CENTER);
		style222.setAlignment(XSSFCellStyle.ALIGN_CENTER);
		style222.setWrapText(true);
		style222.setFont(font222);

		Font font3 = book.createFont();
		font3.setColor((short) 12);// ????????
		font3.setFontHeightInPoints((short) 14);
		font3.setBoldweight(Font.BOLDWEIGHT_BOLD); // ????????
		XSSFCellStyle style3 = book.createCellStyle();
		style3.setBorderBottom(CellStyle.BORDER_MEDIUM); // ????????
		style3.setBorderLeft(CellStyle.BORDER_MEDIUM); // ????????
		style3.setBorderRight(CellStyle.BORDER_MEDIUM); // ????????
		style3.setBorderTop(CellStyle.BORDER_MEDIUM); // ????????
		style3.setVerticalAlignment(XSSFCellStyle.VERTICAL_CENTER);
		style3.setAlignment(XSSFCellStyle.ALIGN_CENTER);
		style3.setFont(font3);

		XSSFCellStyle style4 = book.createCellStyle();
		style4.setBorderBottom(XSSFCellStyle.BORDER_THIN); // ??????
		// style4.setBorderLeft(XSSFCellStyle.BORDER_THIN);// ??????
		style4.setBorderTop(XSSFCellStyle.BORDER_THIN);// ??????
		style4.setBorderRight(XSSFCellStyle.BORDER_THIN);// ??????
		style4.setVerticalAlignment(XSSFCellStyle.VERTICAL_CENTER);
		style4.setAlignment(XSSFCellStyle.ALIGN_CENTER);
		style4.setFont(font);

		XSSFCellStyle style5 = book.createCellStyle();
		style5.setBorderBottom(XSSFCellStyle.BORDER_THIN); // ??????
		style5.setBorderLeft(XSSFCellStyle.BORDER_THIN);// ??????
		style5.setBorderTop(XSSFCellStyle.BORDER_THIN);// ??????
		// style5.setBorderRight(XSSFCellStyle.BORDER_THIN);// ??????
		style5.setVerticalAlignment(XSSFCellStyle.VERTICAL_CENTER);
		style5.setAlignment(XSSFCellStyle.ALIGN_CENTER);
		style5.setFont(font);

		// ????????????
		Font font4 = book.createFont();
		font4.setColor((short) 2);// ????????
		font4.setFontHeightInPoints((short) 10);

		XSSFCellStyle style44 = book.createCellStyle();
		style44.setBorderBottom(XSSFCellStyle.BORDER_THIN); // ??????
		// style4.setBorderLeft(XSSFCellStyle.BORDER_THIN);// ??????
		style44.setBorderTop(XSSFCellStyle.BORDER_THIN);// ??????
		style44.setBorderRight(XSSFCellStyle.BORDER_THIN);// ??????
		style44.setVerticalAlignment(XSSFCellStyle.VERTICAL_CENTER);
		style44.setAlignment(XSSFCellStyle.ALIGN_CENTER);
		style44.setFont(font4);

		XSSFCellStyle style6 = book.createCellStyle();
		style6.setBorderBottom(XSSFCellStyle.BORDER_NONE); // ??????
		style6.setBorderLeft(XSSFCellStyle.BORDER_MEDIUM);// ??????
		style6.setBorderTop(XSSFCellStyle.BORDER_NONE);// ??????
		style6.setBorderRight(XSSFCellStyle.BORDER_NONE);// ??????

		XSSFCellStyle style7 = book.createCellStyle();
		style7.setBorderBottom(XSSFCellStyle.BORDER_NONE); // ??????
		style7.setBorderLeft(XSSFCellStyle.BORDER_NONE);// ??????
		style7.setBorderTop(XSSFCellStyle.BORDER_NONE);// ??????
		style7.setBorderRight(XSSFCellStyle.BORDER_NONE);// ??????

		// ??????????
		Font fontpink = book.createFont();
		fontpink.setColor((short) 12);// ????????
		fontpink.setFontName("MS PGothic");
		fontpink.setFontHeightInPoints((short) 9);

		XSSFCellStyle stylepink = book.createCellStyle();
		stylepink.setFillForegroundColor(IndexedColors.ROSE.getIndex());
		stylepink.setFillPattern(CellStyle.SOLID_FOREGROUND);
		stylepink.setBorderBottom(XSSFCellStyle.BORDER_THIN); // ??????
		stylepink.setBorderLeft(XSSFCellStyle.BORDER_THIN);// ??????
		stylepink.setBorderTop(XSSFCellStyle.BORDER_THIN);// ??????
		stylepink.setBorderRight(XSSFCellStyle.BORDER_THIN);// ??????
		stylepink.setVerticalAlignment(XSSFCellStyle.VERTICAL_CENTER);
		stylepink.setAlignment(XSSFCellStyle.ALIGN_CENTER);
		stylepink.setFont(fontpink);

		if (updateflag) {
			int number = 0;
			for (int i = 0; i < sheetnum; i++) {
				String sheetname = book.getSheetName(i);
				if (sheetname.contains("PSW")) {
					if (sheetAtIndex <= i) {
						number++;
					}
				}
			}
			index = index + page - 1;
			if (number < page) {
				if (page - number > 0) {
					for (int i = 0; i < page - number; i++) {
						XSSFSheet newsheet = book.cloneSheet(sheetAtIndex);
						book.setSheetOrder(newsheet.getSheetName(), index);
						index++;
					}
				}
			}
		} else {
			if (page > 1) {
				for (int i = 1; i < page; i++) {
					XSSFSheet newsheet = book.cloneSheet(sheetAtIndex);
					book.setSheetOrder(newsheet.getSheetName(), index);
					index++;
				}
			}
		}
		System.out.println("page: " + page);

		int shnum = 0;

		int maxRepressure = 0;// ????????????
		int minRepressure = 99999999;// ????????????
		double sumrevalue = 0;// ????????

		int datanum = 0; // ??????????????????
		
		boolean isCucalPara = true; //????????????????????

		for (int i = sheetAtIndex; i < index; i++) {
			XSSFSheet sheet = book.getSheetAt(i);

			setStringCellAndStyle(sheet, TransformerNumber, 5, 8, style2, Cell.CELL_TYPE_STRING);// ??????????
			setStringCellAndStyle(sheet, Guncode, 5, 19, style3, Cell.CELL_TYPE_STRING);// ????????
			if (!sopflag) {
				setStringCellAndStyle(sheet, ElectrodeVol, 7, 108, style2, Cell.CELL_TYPE_STRING);// ????????
			}

			if (i == index - 1) {
				for (int j = 0; j + 36 * shnum < hdinfo.size() + symmetryhdinfo.size(); j++) {
					WeldPointBoardInformation wpb = new WeldPointBoardInformation();
					if (j + 36 * shnum > hdinfo.size() - 1) {
						wpb = symmetryhdinfo.get(j + 36 * shnum - hdinfo.size());
					} else {
						wpb = hdinfo.get(j + 36 * shnum);
					}

					String weldno = wpb.getWeldno(); // ????????
					String pageNo = "";// ????
					String dot = "";// ??????
					// ??????????????????????????
					if (notelist.containsKey(weldno)) {
						String[] pd = notelist.get(weldno);
						pageNo = pd[0];
						dot = pd[1];
					}
					String importance = wpb.getImportance(); // ??????
					// ??????????????????????????????
					if (importance != null && !importance.isEmpty()) {
						if (!Import.contains(importance)) {
							Import.add(importance);
						}
					}
					String boardnumber1 = wpb.getBoardnumber1(); // ????1????
					String boardname1 = wpb.getBoardname1(); // ????1????
					String partmaterial1 = wpb.getPartmaterial1(); // ????1????
					String partthickness1 = wpb.getPartthickness1(); // ????1????
					String boardnumber2 = wpb.getBoardnumber2(); // ????2????
					String boardname2 = wpb.getBoardname2(); // ????2????
					String partmaterial2 = wpb.getPartmaterial2(); // ????2????
					String partthickness2 = wpb.getPartthickness2(); // ????2????
					String boardnumber3 = wpb.getBoardnumber3(); // ????3????
					String boardname3 = wpb.getBoardname3(); // ????3????
					String partmaterial3 = wpb.getPartmaterial3(); // ????3????
					String partthickness3 = wpb.getPartthickness3(); // ????3????
					String layersnum = wpb.getLayersnum(); // ??????
					String gagi = wpb.getGagi(); // GA /GI
					String sheetstrength440 = wpb.getSheetstrength440(); // ????????(Mpa)440
					String sheetstrength590 = wpb.getSheetstrength590(); // ????????(Mpa)590
					String sheetstrength = wpb.getSheetstrength(); // ????????(Mpa)>590
					String basethickness = wpb.getBasethickness(); // ????????
					String sheetstrength12 = wpb.getSheetstrength12(); // ????????(Mpa)1.2G

					String poweroncurent2 = "";// ????????????
					String RecomWeldForce = "";// ???? ??????(N)
					String CurrentSerie = "";// ????????

					// ????????????????????????????????????????????
					boolean partmaterialFlag1 = true;
					boolean partmaterialFlag2 = true;
					boolean partmaterialFlag3 = true;
					String gagi1 = wpb.getGagi1();								
					String gagi2 = wpb.getGagi2();								
					String gagi3 = wpb.getGagi3();

					// ??????????????????GA/GI????
					if (MaterialMap != null) {
						for (Map.Entry<String, List<String>> entry : MaterialMap.entrySet()) {
							String MaterialNo = entry.getKey();
							List<String> infolist = entry.getValue();
							System.out.println(MaterialNo + infolist);
							if(!"GA".equals(gagi1) && !"GI".equals(gagi1))
							{
								if (Util.getIsEqueal(partmaterial1, MaterialNo)) {
									if ("??".equalsIgnoreCase(infolist.get(1))) {
										partmaterialFlag1 = false;
										isCucalPara = false;
									}
								}
							}
							if(!"GA".equals(gagi2) && !"GI".equals(gagi2))
							{
								if (Util.getIsEqueal(partmaterial2, MaterialNo)) {
									if ("??".equalsIgnoreCase(infolist.get(1))) {
										partmaterialFlag2 = false;
										isCucalPara = false;
									}
								}
							}
							if(!"GA".equals(gagi3) && !"GI".equals(gagi3))
							{
								if (Util.getIsEqueal(partmaterial3, MaterialNo)) {
									if ("??".equalsIgnoreCase(infolist.get(1))) {
										partmaterialFlag3 = false;
										isCucalPara = false;
									}
								}
							}																				
						}
					}
					System.out.println(partmaterialFlag1);
					System.out.println(partmaterialFlag2);
					System.out.println(partmaterialFlag3);
					System.out.println(isCucalPara);
					// ????????????????????
					if (partmaterialFlag1 && partmaterialFlag2 && partmaterialFlag3) {
						if (hjmap.containsKey(weldno)) {
							String[] curenre = hjmap.get(weldno);
							RecomWeldForce = curenre[4];
							poweroncurent2 = curenre[3];
							CurrentSerie = curenre[1];
						}
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
							datanum++;
						}
					}

					setStringCellAndStyle(sheet, pageNo, 11 + j, 2, style, Cell.CELL_TYPE_STRING);
					setStringCellAndStyle(sheet, dot, 11 + j, 6, style5, Cell.CELL_TYPE_STRING);
					setStringCellAndStyle(sheet, importance, 11 + j, 4, style, Cell.CELL_TYPE_STRING);
					if (updateflag && pswlist != null) {
						if (!pswlist.contains(weldno)) {
							XSSFCellStyle newstyle = getXSSFStyle(book, sheet, 11 + j, 8, 2, -1);
							setStringCellAndStyle2(sheet, weldno, 11 + j, 8, newstyle, Cell.CELL_TYPE_STRING);
						} else {
							XSSFCellStyle newstyle = getXSSFStyle(book, sheet, 11 + j, 8, 12, -1);
							setStringCellAndStyle2(sheet, weldno, 11 + j, 8, newstyle, Cell.CELL_TYPE_STRING);
						}
					} else {
						XSSFCellStyle newstyle = getXSSFStyle(book, sheet, 11 + j, 8, 12, -1);
						setStringCellAndStyle2(sheet, weldno, 11 + j, 8, newstyle, Cell.CELL_TYPE_STRING);
					}
					setStringCellAndStyle(sheet, boardnumber1, 11 + j, 13, style, Cell.CELL_TYPE_STRING);
					setStringCellAndStyle(sheet, boardname1, 11 + j, 16, style, Cell.CELL_TYPE_STRING);
//					if (!partmaterialFlag1) {
//						setStringCellAndStyle2(sheet, partmaterial1, 11 + j, 29, stylepink, Cell.CELL_TYPE_STRING);
//					} else {
//						setStringCellAndStyle2(sheet, partmaterial1, 11 + j, 29, style, Cell.CELL_TYPE_STRING);
//					}
					setStringCellAndStyle(sheet, partmaterial1, 11 + j, 29, style, Cell.CELL_TYPE_STRING);
					
					setStringCellAndStyle(sheet, partthickness1, 11 + j, 36, style, 11);
					setStringCellAndStyle(sheet, boardnumber2, 11 + j, 39, style, Cell.CELL_TYPE_STRING);
					setStringCellAndStyle(sheet, boardname2, 11 + j, 42, style, Cell.CELL_TYPE_STRING);
//					if (!partmaterialFlag2) {
//						setStringCellAndStyle2(sheet, partmaterial2, 11 + j, 55, stylepink, Cell.CELL_TYPE_STRING);
//					} else {
//						setStringCellAndStyle2(sheet, partmaterial2, 11 + j, 55, style, Cell.CELL_TYPE_STRING);
//					}
					setStringCellAndStyle(sheet, partmaterial2, 11 + j, 55, style, Cell.CELL_TYPE_STRING);

					setStringCellAndStyle(sheet, partthickness2, 11 + j, 62, style, 11);
					setStringCellAndStyle(sheet, boardnumber3, 11 + j, 65, style, Cell.CELL_TYPE_STRING);
					setStringCellAndStyle(sheet, boardname3, 11 + j, 68, style, Cell.CELL_TYPE_STRING);
//					if (!partmaterialFlag3) {
//						setStringCellAndStyle2(sheet, partmaterial3, 11 + j, 81, stylepink, Cell.CELL_TYPE_STRING);
//					} else {
//						setStringCellAndStyle2(sheet, partmaterial3, 11 + j, 81, style, Cell.CELL_TYPE_STRING);
//					}
					setStringCellAndStyle(sheet, partmaterial3, 11 + j, 81, style, Cell.CELL_TYPE_STRING);
					
					setStringCellAndStyle(sheet, partthickness3, 11 + j, 88, style, 11);
					setStringCellAndStyle(sheet, layersnum, 11 + j, 91, style, 10);
					if (gagi != null && !gagi.isEmpty()) {
						setStringCellAndStyle(sheet, gagi, 11 + j, 93, style, Cell.CELL_TYPE_STRING);
					} else {
						setStringCellAndStyle(sheet, "-", 11 + j, 93, style, Cell.CELL_TYPE_STRING);
					}
					setStringCellAndStyle(sheet, sheetstrength440, 11 + j, 95, style, 10);
					setStringCellAndStyle(sheet, sheetstrength590, 11 + j, 97, style, 10);
					setStringCellAndStyle(sheet, sheetstrength, 11 + j, 99, style, 10);
					setStringCellAndStyle(sheet, basethickness, 11 + j, 102, style, 11);

					// ??????1.2g????????????????????????????
					if (sheetstrength12.equals("1.2g")) {
						setStringCellAndStyle(sheet, "", 11 + j, 105, style, 10);
						setStringCellAndStyle(sheet, "", 11 + j, 108, style, 10);
						setStringCellAndStyle(sheet, "", 11 + j, 111, style, 10);
					} else {

						// ????????????
						if (Util.isNumber(poweroncurent2)) {
							int curent = 0;
							curent = (int) (Double.parseDouble(poweroncurent2) * 1000);
							poweroncurent2 = Integer.toString(curent);
						}
						setStringCellAndStyle(sheet, CurrentSerie, 11 + j, 105, style, 10);
						setStringCellAndStyle(sheet, RecomWeldForce, 11 + j, 108, style, 10);
						setStringCellAndStyle(sheet, poweroncurent2, 11 + j, 111, style, 10);
					}

				}
			} else {
				for (int j = 0; j + 36 * shnum < 36 + 36 * shnum; j++) {
//					WeldPointBoardInformation wpb = hdinfo.get(j + 36 * shnum);
					WeldPointBoardInformation wpb = new WeldPointBoardInformation();
					if (j + 36 * shnum > hdinfo.size() - 1) {
						wpb = symmetryhdinfo.get(j + 36 * shnum - hdinfo.size());
					} else {
						wpb = hdinfo.get(j + 36 * shnum);
					}

					String weldno = wpb.getWeldno(); // ????????
					String pageNo = "";// ????
					String dot = "";// ??????
					// ??????????????????????????
					if (notelist.containsKey(weldno)) {
						String[] pd = notelist.get(weldno);
						pageNo = pd[0];
						dot = pd[1];
					}
					String importance = wpb.getImportance(); // ??????
					// ??????????????????????????????
					if (importance != null && !importance.isEmpty()) {
						if (!Import.contains(importance)) {
							Import.add(importance);
						}
					}
					String boardnumber1 = wpb.getBoardnumber1(); // ????1????
					String boardname1 = wpb.getBoardname1(); // ????1????
					String partmaterial1 = wpb.getPartmaterial1(); // ????1????
					String partthickness1 = wpb.getPartthickness1(); // ????1????
					String boardnumber2 = wpb.getBoardnumber2(); // ????2????
					String boardname2 = wpb.getBoardname2(); // ????2????
					String partmaterial2 = wpb.getPartmaterial2(); // ????2????
					String partthickness2 = wpb.getPartthickness2(); // ????2????
					String boardnumber3 = wpb.getBoardnumber3(); // ????3????
					String boardname3 = wpb.getBoardname3(); // ????3????
					String partmaterial3 = wpb.getPartmaterial3(); // ????3????
					String partthickness3 = wpb.getPartthickness3(); // ????3????
					String layersnum = wpb.getLayersnum(); // ??????
					String gagi = wpb.getGagi(); // GA /GI
					String sheetstrength440 = wpb.getSheetstrength440(); // ????????(Mpa)440
					String sheetstrength590 = wpb.getSheetstrength590(); // ????????(Mpa)590
					String sheetstrength = wpb.getSheetstrength(); // ????????(Mpa)>590
					String basethickness = wpb.getBasethickness(); // ????????
					String sheetstrength12 = wpb.getSheetstrength12(); // ????????(Mpa)1.2G

					String poweroncurent2 = "";// ????????????
					String RecomWeldForce = "";// ???? ??????(N)
					String CurrentSerie = "";// ????????
					// ????????????????????????????????????????????
					boolean partmaterialFlag1 = true;
					boolean partmaterialFlag2 = true;
					boolean partmaterialFlag3 = true;
					String gagi1 = wpb.getGagi1();								
					String gagi2 = wpb.getGagi2();								
					String gagi3 = wpb.getGagi3();

					// ??????????????????GA/GI????
					if (MaterialMap != null) {
						for (Map.Entry<String, List<String>> entry : MaterialMap.entrySet()) {
							String MaterialNo = entry.getKey();
							List<String> infolist = entry.getValue();
							if(!"GA".equals(gagi1) && !"GI".equals(gagi1))
							{
								if (Util.getIsEqueal(partmaterial1, MaterialNo)) {
									if ("??".equalsIgnoreCase(infolist.get(1))) {
										partmaterialFlag1 = false;
										isCucalPara = false;
									}
								}
							}
							if(!"GA".equals(gagi2) && !"GI".equals(gagi2))
							{
								if (Util.getIsEqueal(partmaterial2, MaterialNo)) {
									if ("??".equalsIgnoreCase(infolist.get(1))) {
										partmaterialFlag2 = false;
										isCucalPara = false;
									}
								}
							}
							if(!"GA".equals(gagi3) && !"GI".equals(gagi3))
							{
								if (Util.getIsEqueal(partmaterial3, MaterialNo)) {
									if ("??".equalsIgnoreCase(infolist.get(1))) {
										partmaterialFlag3 = false;
										isCucalPara = false;
									}
								}
							}																				
						}
					}
					System.out.println(partmaterialFlag1);
					System.out.println(partmaterialFlag2);
					System.out.println(partmaterialFlag3);
					System.out.println(isCucalPara);
					// ????????????????????
					if (partmaterialFlag1 && partmaterialFlag2 && partmaterialFlag3) {
						if (hjmap.containsKey(weldno)) {
							String[] curenre = hjmap.get(weldno);
							RecomWeldForce = curenre[4];
							poweroncurent2 = curenre[3];
							CurrentSerie = curenre[1];
						}
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
							datanum++;
						}
					}

					setStringCellAndStyle(sheet, pageNo, 11 + j, 2, style, Cell.CELL_TYPE_STRING);
					setStringCellAndStyle(sheet, dot, 11 + j, 6, style5, Cell.CELL_TYPE_STRING);
					setStringCellAndStyle(sheet, importance, 11 + j, 4, style, Cell.CELL_TYPE_STRING);
					if (updateflag && pswlist != null) {
						if (!pswlist.contains(weldno)) {
							XSSFCellStyle newstyle = getXSSFStyle(book, sheet, 11 + j, 8, 2, -1);
							setStringCellAndStyle2(sheet, weldno, 11 + j, 8, newstyle, Cell.CELL_TYPE_STRING);
						} else {
							XSSFCellStyle newstyle = getXSSFStyle(book, sheet, 11 + j, 8, 12, -1);
							setStringCellAndStyle2(sheet, weldno, 11 + j, 8, newstyle, Cell.CELL_TYPE_STRING);
						}
					} else {
						XSSFCellStyle newstyle = getXSSFStyle(book, sheet, 11 + j, 8, 12, -1);
						setStringCellAndStyle2(sheet, weldno, 11 + j, 8, newstyle, Cell.CELL_TYPE_STRING);
					}
					setStringCellAndStyle(sheet, boardnumber1, 11 + j, 13, style, Cell.CELL_TYPE_STRING);
					setStringCellAndStyle(sheet, boardname1, 11 + j, 16, style, Cell.CELL_TYPE_STRING);
//					if (!partmaterialFlag1) {
//						setStringCellAndStyle2(sheet, partmaterial1, 11 + j, 29, stylepink, Cell.CELL_TYPE_STRING);
//					} else {
//						setStringCellAndStyle2(sheet, partmaterial1, 11 + j, 29, style, Cell.CELL_TYPE_STRING);
//					}
					setStringCellAndStyle(sheet, partmaterial1, 11 + j, 29, style, Cell.CELL_TYPE_STRING);
					
					setStringCellAndStyle(sheet, partthickness1, 11 + j, 36, style, 11);
					setStringCellAndStyle(sheet, boardnumber2, 11 + j, 39, style, Cell.CELL_TYPE_STRING);
					setStringCellAndStyle(sheet, boardname2, 11 + j, 42, style, Cell.CELL_TYPE_STRING);
//					if (!partmaterialFlag2) {
//						setStringCellAndStyle2(sheet, partmaterial2, 11 + j, 55, stylepink, Cell.CELL_TYPE_STRING);
//					} else {
//						setStringCellAndStyle2(sheet, partmaterial2, 11 + j, 55, style, Cell.CELL_TYPE_STRING);
//					}
					setStringCellAndStyle(sheet, partmaterial2, 11 + j, 55, style, Cell.CELL_TYPE_STRING);

					setStringCellAndStyle(sheet, partthickness2, 11 + j, 62, style, 11);
					setStringCellAndStyle(sheet, boardnumber3, 11 + j, 65, style, Cell.CELL_TYPE_STRING);
					setStringCellAndStyle(sheet, boardname3, 11 + j, 68, style, Cell.CELL_TYPE_STRING);
//					if (!partmaterialFlag3) {
//						setStringCellAndStyle2(sheet, partmaterial3, 11 + j, 81, stylepink, Cell.CELL_TYPE_STRING);
//					} else {
//						setStringCellAndStyle2(sheet, partmaterial3, 11 + j, 81, style, Cell.CELL_TYPE_STRING);
//					}
					setStringCellAndStyle(sheet, partmaterial3, 11 + j, 81, style, Cell.CELL_TYPE_STRING);
					
					setStringCellAndStyle(sheet, partthickness3, 11 + j, 88, style, 11);
					setStringCellAndStyle(sheet, layersnum, 11 + j, 91, style, 10);
					if (gagi != null && !gagi.isEmpty()) {
						setStringCellAndStyle(sheet, gagi, 11 + j, 93, style, Cell.CELL_TYPE_STRING);
					} else {
						setStringCellAndStyle(sheet, "-", 11 + j, 93, style, Cell.CELL_TYPE_STRING);
					}
					setStringCellAndStyle(sheet, sheetstrength440, 11 + j, 95, style, 10);
					setStringCellAndStyle(sheet, sheetstrength590, 11 + j, 97, style, 10);
					setStringCellAndStyle(sheet, sheetstrength, 11 + j, 99, style, 10);
					setStringCellAndStyle(sheet, basethickness, 11 + j, 102, style, 11);

					// ??????1.2g????????????????????????????
					if (sheetstrength12.equals("1.2g")) {
						setStringCellAndStyle(sheet, "", 11 + j, 105, style, 10);
						setStringCellAndStyle(sheet, "", 11 + j, 108, style, 10);
						setStringCellAndStyle(sheet, "", 11 + j, 111, style, 10);
					} else {
						// ????????????
						if (Util.isNumber(poweroncurent2)) {
							int curent = 0;
							curent = (int) (Double.parseDouble(poweroncurent2) * 1000);
							poweroncurent2 = Integer.toString(curent);
						}
						setStringCellAndStyle(sheet, CurrentSerie, 11 + j, 105, style, 10);
						setStringCellAndStyle(sheet, RecomWeldForce, 11 + j, 108, style, 10);
						setStringCellAndStyle(sheet, poweroncurent2, 11 + j, 111, style, 10);
					}
				}
			}
//			// ???????? ??????????
//			XSSFRow row = sheet.getRow(5);
//			if (row != null) {
//				XSSFCell cell = row.getCell(19);
//				if (cell != null) {
//					NewOutputDataToExcel.setFontSize(book, cell, (short) 16);
//				}
//			}
			shnum++;
		}

		// ????????
		String[] tatolcurenre = new String[12];
		// ????????????????????????????????
		if(!isCucalPara || datanum == 0)
		{			
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
		}
		else
		{
			tatolcurenre = getAverageParameterValues(maxRepressure, minRepressure, sumrevalue, datanum);
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
		

		// ????????????????????
		for (int i = sheetAtIndex; i < index; i++) {
			XSSFSheet sheet = book.getSheetAt(i);
			if (!sopflag) {
				for (int j = 0; j < tatolcurenre.length; j++) {
					if (j == 0) {
						setStringCellAndStyle(sheet, tatolcurenre[j], 7, 36 + j * 6, style20, Cell.CELL_TYPE_STRING);// ??????~????
					} else if (j == tatolcurenre.length - 1) {
						setStringCellAndStyle(sheet, tatolcurenre[j], 7, 36 + j * 6, style21, Cell.CELL_TYPE_STRING);// ??????~????
					} else {
						setStringCellAndStyle(sheet, tatolcurenre[j], 7, 36 + j * 6, style22, Cell.CELL_TYPE_STRING);// ??????~????
					}

				}
			} else {
				if (!updateflag) {
					// ????????????
					for (int j = 0; j < 3; j++) {
						for (int k = 0; k < 77; k++) {
							if (k == 0) {
								setStringCellAndStyle2(sheet, "", 5 + j, 36 + k, style6, Cell.CELL_TYPE_STRING);// ??????~????
							} else {
								setStringCellAndStyle2(sheet, "", 5 + j, 36 + k, style7, Cell.CELL_TYPE_STRING);// ??????~????
							}
						}
					}
				}
			}
			// ??????????????????????????
			if (updateflag) {
				// ??????????????????????????????????????????
				XSSFRow terow = sheet.getRow(48);
				XSSFCell tecell = terow.getCell(108);
				String preedtion = tecell.getStringCellValue();
				boolean teflag = getIsTeSOPAfter(preedtion);
				if (!teflag) // ????????
				{
					setStringCellAndStyle(sheet, "??????", 5, 36, style202, Cell.CELL_TYPE_STRING);// ??????~????
					setStringCellAndStyle(sheet, "????????", 5, 42, style222, Cell.CELL_TYPE_STRING);// ??????~????
					setStringCellAndStyle(sheet, "????????", 5, 48, style222, Cell.CELL_TYPE_STRING);// ??????~????
					setStringCellAndStyle(sheet, "????          ????????", 5, 54, style222, Cell.CELL_TYPE_STRING);// ??????~????
					setStringCellAndStyle(sheet, "????          ????????", 5, 60, style222, Cell.CELL_TYPE_STRING);// ??????~????
					setStringCellAndStyle(sheet, "??????????", 5, 66, style222, Cell.CELL_TYPE_STRING);// ??????~????
					setStringCellAndStyle(sheet, "????          ????????", 5, 72, style222, Cell.CELL_TYPE_STRING);// ??????~????
					setStringCellAndStyle(sheet, "????          ????????", 5, 78, style222, Cell.CELL_TYPE_STRING);// ??????~????
					setStringCellAndStyle(sheet, "??????????", 5, 84, style222, Cell.CELL_TYPE_STRING);// ??????~????
					setStringCellAndStyle(sheet, "????          ????????", 5, 90, style222, Cell.CELL_TYPE_STRING);// ??????~????
					setStringCellAndStyle(sheet, "????         ????????", 5, 96, style222, Cell.CELL_TYPE_STRING);// ??????~????
					setStringCellAndStyle(sheet, "????", 5, 102, style212, Cell.CELL_TYPE_STRING);// ??????~????
					setStringCellAndStyle(sheet, "????????????", 5, 108, style212, Cell.CELL_TYPE_STRING);// ??????~????
																									// ElectrodeVol
					setStringCellAndStyle(sheet, ElectrodeVol, 7, 108, style212, Cell.CELL_TYPE_STRING);//

					for (int j = 0; j < tatolcurenre.length; j++) {
						if (j == 0) {
							setStringCellAndStyle(sheet, tatolcurenre[j], 7, 36 + j * 6, style20,
									Cell.CELL_TYPE_STRING);// ??????~????
						} else if (j == tatolcurenre.length - 1) {
							setStringCellAndStyle(sheet, tatolcurenre[j], 7, 36 + j * 6, style21,
									Cell.CELL_TYPE_STRING);// ??????~????
						} else {
							setStringCellAndStyle(sheet, tatolcurenre[j], 7, 36 + j * 6, style22,
									Cell.CELL_TYPE_STRING);// ??????~????
						}
					}
				}
			}

		}

		String[] provalues = new String[11];
		int indexCount = 0;
		for (int i = 0; i < tatolcurenre.length; i++) {
			if (i != 1) {
				provalues[indexCount] = tatolcurenre[i].replace("c.", "").replace("KA", "").replace("N", "");
				indexCount++;
			}
		}
		TCComponentItemRevision rev = bl.getItemRevision();

		// ????????????????????????????????
		String[] properties = { "b8_WeldForce", "b8_RiseTime", "b8_CurrentTime1", "b8_Current1", "b8_Cool1",
				"b8_CurrentTime2", "b8_Current2", "b8_Cool2", "b8_CurrentTime3", "b8_Current3", "b8_KeepTime", };
		// ????????
		TCProperty[] pp = rev.getTCProperties(properties);
		if (pp != null && pp[0] != null) {
			rev.setProperties(properties, provalues);
			rev.lock();
			rev.save();
			rev.unlock();
		}

		return error;
	}

	private void getSymmetryHanJieParater(ArrayList hdlist, Map<String, String[]> hjmap) throws TCException {
		// TODO Auto-generated method stub
		if (hdlist != null && hdlist.size() > 0) {
			TCComponent[] tccs = new TCComponent[hdlist.size()];
			for (int i = 0; i < hdlist.size(); i++) {
				TCComponentBOMLine bl = (TCComponentBOMLine) hdlist.get(i);
				TCComponentItemRevision rev = bl.getItemRevision();
				tccs[i] = rev;
				String weldno = Util.getProperty(rev, "object_name");
			}
			String[] properties = { "object_name", "b8_CurrentSerie_Nissan", "b8_CurrentSerie_DFL", "b8_Current2",
					"b8_RecomWeldForce" };
			String[][] values = Util.getAllProperties(session, tccs, properties);
			for (int j = 0; j < values.length; j++) {
				String[] proper = values[j];
				hjmap.put(proper[0], proper);
			}
		}
		return;
	}

	// ????????????
	private String[] getAverageParameterValues(int maxRepressure, int minRepressure, double sumrevalue, int size) {
		// TODO Auto-generated method stub
		String[] values = new String[12];
		String press = "";// ??????
		String Preloadingtime = "15c.";// ???????? ??????????
		String uptime = "";// ????????
		String powerontime1 = "";// ???? ????????
		String poweroncurent1 = "";// ???? ????????
		String coolingtime1 = "";// ??????????
		String powerontime2 = "";// ????????????
		String poweroncurent2 = "";// ????????????
		String coolingtime2 = "";// ??????????
		String powerontime3 = "";// ???? ????????
		String poweroncurent3 = "";// ???? ????????
		String maintain = "";// ????

		int prepress = (maxRepressure + minRepressure) / 2;
		press = Integer.toString(prepress);
		// ????????????
//		Object[] obj = baseinfoExcelReader.getCalculationParameter(app, "B8_Calculation_Parameter_Name");
//		List<CurrentandVoltage> cv = new ArrayList<CurrentandVoltage>();
//		if (obj != null) {
//			if (obj[1] != null) {
//				cv = (List<CurrentandVoltage>) obj[1];
//			} else {
//				System.out.println("????????24?????????????????? ??????????????");
//			}
//		}
		// ??????????
		BigDecimal biga1 = new BigDecimal(Double.toString(sumrevalue));
		BigDecimal bigsize = new BigDecimal(Double.toString(size));
		double average = biga1.divide(bigsize, 8, BigDecimal.ROUND_HALF_UP).doubleValue();
		// 255?????????????????? ????????
		CurrentandVoltage currentandVoltage = getCurrentandVoltage(average, cv);
		if (currentandVoltage != null) {
			uptime = currentandVoltage.getBvalue() + "c.";// ????????
			powerontime1 = currentandVoltage.getCvalue() + "c.";// ???? ????????
			poweroncurent1 = currentandVoltage.getEvalue() + "KA";// ???? ????????
			coolingtime1 = currentandVoltage.getFvalue() + "c.";// ??????????
			powerontime2 = currentandVoltage.getGvalue() + "c.";// ????????????
			poweroncurent2 = currentandVoltage.getIvalue() + "KA";// ????????????
			coolingtime2 = currentandVoltage.getJvalue() + "c.";// ??????????
			powerontime3 = currentandVoltage.getKvalue() + "c.";// ???? ????????
			poweroncurent3 = currentandVoltage.getMvalue() + "KA";// ???? ????????
			maintain = currentandVoltage.getNvalue() + "c.";// ????;
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

	// 255?????????????????? ????????
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
	 * RSW????????????
	 */
	private void RSWServoinformationProcessing(XSSFWorkbook book, TCComponentBOMLine bl, TCComponentBOMLine gwbl,
			String name, Map<String, TCComponentBOMLine> symmetrymap) throws TCException {
		// TODO Auto-generated method stub
		// ????RSW????sheet
		int sheetnum = 0;
		sheetnum = book.getNumberOfSheets();
		int sheetAtIndex = -1; // RSW????????????
		for (int i = 0; i < sheetnum; i++) {
			String sheetname = book.getSheetName(i);
			if (sheetname.contains(name)) {
				sheetAtIndex = i;
				break;
			}
		}
		if (sheetAtIndex == -1) {
			return;
		}
		// ??????????????????????
		ArrayList gunlist = new ArrayList();
		ArrayList hdlist = new ArrayList();
		// ArrayList robotlist = new ArrayList();
		String guntypename = Util.getObjectDisplayName(session, "B8_BIWGun");
		String[] propertys = new String[] { "bl_item_object_type", "bl_item_object_type" };
		String[] values = new String[] { guntypename, guntypename };
		String weldtypename = Util.getObjectDisplayName(session, "WeldPoint");
		String[] propertys2 = new String[] { "bl_item_object_type", "bl_item_object_type" };
		String[] values2 = new String[] { weldtypename, weldtypename };

//		String[] propertys3 = new String[] { "bl_item_object_type", "bl_item_object_type" };
//		String[] values3 = new String[] { "??????", "BIW Robot" };

		gunlist = Util.searchBOMLine(bl, "OR", propertys, "==", values);
		hdlist = Util.searchBOMLine(bl, "OR", propertys2, "==", values2);
		// robotlist = Util.searchBOMLine(bl, "OR", propertys3, "==", values3);
		// ??????????????SOP????????????SOP??????????????????
		boolean sopflag = getIsSOPAfter();
		String Discretename = Util.getProperty(bl, "bl_rev_object_name");
		// ????????????????????????
		String stationname = Util.getProperty(gwbl, "bl_rev_object_name");// ????
		String robotname = "";// ??????
		String Guncode = "";// ????????
		String ElectrodeVol = "";// ???????? bl_B8_BIWGunRevision_b8_ElectrodeVol
		TCComponentItemRevision blrev = null;
		if (gunlist.size() > 0) {
			TCComponentBOMLine gunbl = (TCComponentBOMLine) gunlist.get(0);
			blrev = gunbl.getItemRevision();
			Guncode = Util.getProperty(blrev, "b8_Model");
			ElectrodeVol = Util.getProperty(blrev, "b8_ElectrodeVol");
		}
		robotname = Util.getProperty(bl, "bl_rev_object_name");
//		if (robotlist.size() > 0) {
//			TCComponentBOMLine robotbl = (TCComponentBOMLine) robotlist.get(0);
//			robotname = Util.getProperty(robotbl, "bl_rev_object_name");
//		}
		// ??????????????????????????????????
		String[] nameArr = Discretename.split("\\\\");
		if (nameArr.length > 1) {
			Guncode = nameArr[1];
		} else {
			Guncode = "";
		}
		robotname = nameArr[0];
		
		// ??????????????????????????????????
		List<WeldPointBoardInformation> hdinfo = new ArrayList<WeldPointBoardInformation>();// ????????????

		hdinfo = getBoardInformation(baseinfolist, hdlist);

		Map<String, String[]> hjmap = getHanJieParater(hdlist);

		if (hdinfo == null || hdinfo.size() < 1) {
			System.out.println("????????????????????????????");
			return;
		}
		// ????????????????????
		Comparator comparator = getComParatorByfirstpart();
		Collections.sort(hdinfo, comparator);

		// ??????????????????????
		Comparator comparator1 = getComParatorBySecondpart();
		Collections.sort(hdinfo, comparator1);

		// ????????????????????
		Comparator comparator11 = getComParatorByThistypart();
		Collections.sort(hdinfo, comparator11);

		// ??????????????????????3????????????2????
		Comparator comparator2 = getComParatorBylayersnum();
		Collections.sort(hdinfo, comparator2);

		// ??????????????????????????????????????????????????????????????????????????????????sheet??????
//		String Discretename = Util.getProperty(bl, "bl_rev_object_name");
		// ??????????????????????????????????
		List<WeldPointBoardInformation> symmetryhdinfo = new ArrayList<WeldPointBoardInformation>();// ????????????
		ArrayList symmetryhdlist = new ArrayList();
		if (symmetrymap.containsKey(Discretename)) {
			TCComponentBOMLine symmetrybl = symmetrymap.get(Discretename);
			symmetryhdlist = Util.searchBOMLine(symmetrybl, "OR", propertys2, "==", values2);
			symmetryhdinfo = getBoardInformation(baseinfolist, symmetryhdlist);

			// ????????????????????
			Comparator comparators1 = getComParatorByfirstpart();
			Collections.sort(symmetryhdinfo, comparators1);
			// ??????????????????????
			Comparator comparators2 = getComParatorBySecondpart();
			Collections.sort(symmetryhdinfo, comparators2);

			// ????????????????????
			Comparator comparators3 = getComParatorByThistypart();
			Collections.sort(symmetryhdinfo, comparators3);

			// ??????????????????????3????????????2????
			Comparator comparators4 = getComParatorBylayersnum();
			Collections.sort(symmetryhdinfo, comparators4);

			// ??????????????????
			getSymmetryHanJieParater(symmetryhdlist, hjmap);
		}

		// ????????????????????????,??36????????????
		int hdsum = hdlist.size() + symmetryhdlist.size();
		int page = hdsum / 36 + 1;
		// ??????????????????????sheet????????????????
		if (hdsum % 36 == 0) {
			if (page > 1) {
				page = page - 1;
			}
		}

		// ????page????1????????????sheet??
		int index = sheetAtIndex + 1;

		// ????????????
		Font font = book.createFont();
		font.setColor((short) 12);// ????????
		font.setFontHeightInPoints((short) 9);
		XSSFCellStyle style = book.createCellStyle();
		style.setBorderBottom(XSSFCellStyle.BORDER_THIN); // ??????
		style.setBorderLeft(XSSFCellStyle.BORDER_THIN);// ??????
		style.setBorderTop(XSSFCellStyle.BORDER_THIN);// ??????
		style.setBorderRight(XSSFCellStyle.BORDER_THIN);// ??????
		style.setVerticalAlignment(XSSFCellStyle.VERTICAL_CENTER);
		style.setAlignment(XSSFCellStyle.ALIGN_CENTER);
		style.setFont(font);

		Font font2 = book.createFont();
		font2.setColor((short) 12);// ????????
		font2.setFontHeightInPoints((short) 18);
		font2.setBoldweight(Font.BOLDWEIGHT_BOLD); // ????????
		XSSFCellStyle style2 = book.createCellStyle();
		style2.setBorderBottom(CellStyle.BORDER_MEDIUM); // ????????
		style2.setBorderLeft(CellStyle.BORDER_MEDIUM); // ????????
		style2.setBorderRight(CellStyle.BORDER_MEDIUM); // ????????
		style2.setBorderTop(CellStyle.BORDER_MEDIUM); // ????????
		style2.setVerticalAlignment(XSSFCellStyle.VERTICAL_CENTER);
		style2.setAlignment(XSSFCellStyle.ALIGN_CENTER);
		style2.setFont(font2);

		// ??????????
		XSSFCellStyle style3 = book.createCellStyle();
		style3.setFillForegroundColor(IndexedColors.PINK.getIndex());
		style3.setFillPattern(CellStyle.SOLID_FOREGROUND);
		style3.setBorderBottom(XSSFCellStyle.BORDER_THIN); // ??????
		style3.setBorderLeft(XSSFCellStyle.BORDER_THIN);// ??????
		style3.setBorderTop(XSSFCellStyle.BORDER_THIN);// ??????
		style3.setBorderRight(XSSFCellStyle.BORDER_THIN);// ??????
		style3.setVerticalAlignment(XSSFCellStyle.VERTICAL_CENTER);
		style3.setAlignment(XSSFCellStyle.ALIGN_CENTER);
		style3.setFont(font);
		// ??????????
		Font font3 = book.createFont();
		font3.setColor((short) 1);// ????????
		font3.setFontHeightInPoints((short) 10);
		XSSFCellStyle style4 = book.createCellStyle();
		style4.setFillForegroundColor(IndexedColors.VIOLET.getIndex());
		style4.setFillPattern(CellStyle.SOLID_FOREGROUND);
		style4.setBorderBottom(XSSFCellStyle.BORDER_THIN); // ??????
		style4.setBorderLeft(XSSFCellStyle.BORDER_THIN);// ??????
		style4.setBorderTop(XSSFCellStyle.BORDER_THIN);// ??????
		style4.setBorderRight(XSSFCellStyle.BORDER_THIN);// ??????
		style4.setVerticalAlignment(XSSFCellStyle.VERTICAL_CENTER);
		style4.setAlignment(XSSFCellStyle.ALIGN_CENTER);
		style4.setFont(font3);
		// ??????????
		Font font4 = book.createFont();
		font4.setColor((short) 1);// ????????
		font4.setFontHeightInPoints((short) 10);
		XSSFCellStyle style5 = book.createCellStyle();
		style5.setFillForegroundColor(IndexedColors.SKY_BLUE.getIndex());
		style5.setFillPattern(CellStyle.SOLID_FOREGROUND);
		style5.setBorderBottom(XSSFCellStyle.BORDER_THIN); // ??????
		style5.setBorderLeft(XSSFCellStyle.BORDER_THIN);// ??????
		style5.setBorderTop(XSSFCellStyle.BORDER_THIN);// ??????
		style5.setBorderRight(XSSFCellStyle.BORDER_THIN);// ??????
		style5.setVerticalAlignment(XSSFCellStyle.VERTICAL_CENTER);
		style5.setAlignment(XSSFCellStyle.ALIGN_CENTER);
		style5.setFont(font4);

		XSSFCellStyle style6 = book.createCellStyle();
		style6.setBorderBottom(XSSFCellStyle.BORDER_THIN); // ??????
		// style6.setBorderLeft(XSSFCellStyle.BORDER_THIN);// ??????
		style6.setBorderTop(XSSFCellStyle.BORDER_THIN);// ??????
		style6.setBorderRight(XSSFCellStyle.BORDER_THIN);// ??????
		style6.setVerticalAlignment(XSSFCellStyle.VERTICAL_CENTER);
		style6.setAlignment(XSSFCellStyle.ALIGN_CENTER);
		style6.setFont(font);

		XSSFCellStyle style8 = book.createCellStyle();
		style8.setBorderBottom(XSSFCellStyle.BORDER_THIN); // ??????
		style8.setBorderLeft(XSSFCellStyle.BORDER_THIN);// ??????
		style8.setBorderTop(XSSFCellStyle.BORDER_THIN);// ??????
		// style8.setBorderRight(XSSFCellStyle.BORDER_THIN);// ??????
		style8.setVerticalAlignment(XSSFCellStyle.VERTICAL_CENTER);
		style8.setAlignment(XSSFCellStyle.ALIGN_CENTER);
		style8.setFont(font);

		// ??????????
		Font font5 = book.createFont();
		font4.setFontHeightInPoints((short) 10);
		XSSFCellStyle style7 = book.createCellStyle();
		style7.setFillForegroundColor(IndexedColors.WHITE.getIndex());
		style7.setFillPattern(CellStyle.SOLID_FOREGROUND);
		style7.setBorderBottom(XSSFCellStyle.BORDER_THIN); // ??????
		style7.setBorderLeft(XSSFCellStyle.BORDER_THIN);// ??????
		style7.setBorderTop(XSSFCellStyle.BORDER_THIN);// ??????
		style7.setBorderRight(XSSFCellStyle.BORDER_THIN);// ??????
		style7.setVerticalAlignment(XSSFCellStyle.VERTICAL_CENTER);
		style7.setAlignment(XSSFCellStyle.ALIGN_CENTER);
		style7.setFont(font5);

		// ????????????
		Font font6 = book.createFont();
		font6.setColor((short) 2);// ????????
		font6.setFontHeightInPoints((short) 10);
		XSSFCellStyle style66 = book.createCellStyle();
		style66.setBorderBottom(XSSFCellStyle.BORDER_THIN); // ??????
		// style6.setBorderLeft(XSSFCellStyle.BORDER_THIN);// ??????
		style66.setBorderTop(XSSFCellStyle.BORDER_THIN);// ??????
		style66.setBorderRight(XSSFCellStyle.BORDER_THIN);// ??????
		style66.setVerticalAlignment(XSSFCellStyle.VERTICAL_CENTER);
		style66.setAlignment(XSSFCellStyle.ALIGN_CENTER);
		style66.setFont(font6);

		// ??????????
		Font fontpink = book.createFont();
		fontpink.setColor((short) 12);// ????????
		fontpink.setFontName("MS PGothic");
		fontpink.setFontHeightInPoints((short) 9);

		XSSFCellStyle stylepink = book.createCellStyle();
		stylepink.setFillForegroundColor(IndexedColors.ROSE.getIndex());
		stylepink.setFillPattern(CellStyle.SOLID_FOREGROUND);
		stylepink.setBorderBottom(XSSFCellStyle.BORDER_THIN); // ??????
		stylepink.setBorderLeft(XSSFCellStyle.BORDER_THIN);// ??????
		stylepink.setBorderTop(XSSFCellStyle.BORDER_THIN);// ??????
		stylepink.setBorderRight(XSSFCellStyle.BORDER_THIN);// ??????
		stylepink.setVerticalAlignment(XSSFCellStyle.VERTICAL_CENTER);
		stylepink.setAlignment(XSSFCellStyle.ALIGN_CENTER);
		stylepink.setFont(fontpink);

		if (updateflag) {
			int number = 0;
			for (int i = 0; i < sheetnum; i++) {
				String sheetname = book.getSheetName(i);
				if (sheetname.contains("RSW????")) {
					if (sheetAtIndex <= i) {
						number++;
					}
				}
			}
			index = index + page - 1;
			if (number < page) {
				if (page - number > 0) {
					for (int i = 0; i < page - number; i++) {
						XSSFSheet newsheet = book.cloneSheet(sheetAtIndex);
						book.setSheetOrder(newsheet.getSheetName(), index);
						index++;
					}
				}
			}
		} else {
			if (page > 1) {
				for (int i = 1; i < page; i++) {
					XSSFSheet newsheet = book.cloneSheet(sheetAtIndex);
					book.setSheetOrder(newsheet.getSheetName(), index);
					index++;
				}
			}
		}

		int shnum = 0;

		for (int i = sheetAtIndex; i < index; i++) {
			XSSFSheet sheet = book.getSheetAt(i);

			setStringCellAndStyle(sheet, stationname, 6, 20, style2, Cell.CELL_TYPE_STRING);// ????
			setStringCellAndStyle(sheet, robotname, 6, 31, style2, Cell.CELL_TYPE_STRING);// ??????
			setStringCellAndStyle(sheet, Guncode, 6, 48, style2, Cell.CELL_TYPE_STRING);// ????????
			setStringCellAndStyle(sheet, ElectrodeVol, 6, 65, style2, Cell.CELL_TYPE_STRING);// ????????
																								// bl_B8_BIWGunRevision_b8_ElectrodeVol
			if (i == index - 1) {
				for (int j = 0; j + 36 * shnum < hdinfo.size() + symmetryhdinfo.size(); j++) {
//					WeldPointBoardInformation wpb = hdinfo.get(j + 36 * shnum);
					WeldPointBoardInformation wpb = new WeldPointBoardInformation();
					if (j + 36 * shnum > hdinfo.size() - 1) {
						wpb = symmetryhdinfo.get(j + 36 * shnum - hdinfo.size());
					} else {
						wpb = hdinfo.get(j + 36 * shnum);
					}
					String weldno = wpb.getWeldno(); // ????????
					String pageNo = "";// ????
					String dot = "";// ??????

					// ??????????????????????????
					if (notelist.containsKey(weldno)) {
						String[] pd = notelist.get(weldno);
						pageNo = pd[0];
						dot = pd[1];
					}
					String importance = wpb.getImportance(); // ??????
					// ??????????????????????????????
					if (importance != null && !importance.isEmpty()) {
						if (!Import.contains(importance)) {
							Import.add(importance);
						}
					}
					String boardnumber1 = wpb.getBoardnumber1(); // ????1????
					String boardname1 = wpb.getBoardname1(); // ????1????
					String partmaterial1 = wpb.getPartmaterial1(); // ????1????
					String partthickness1 = wpb.getPartthickness1(); // ????1????
					String boardnumber2 = wpb.getBoardnumber2(); // ????2????
					String boardname2 = wpb.getBoardname2(); // ????2????
					String partmaterial2 = wpb.getPartmaterial2(); // ????2????
					String partthickness2 = wpb.getPartthickness2(); // ????2????
					String boardnumber3 = wpb.getBoardnumber3(); // ????3????
					String boardname3 = wpb.getBoardname3(); // ????3????
					String partmaterial3 = wpb.getPartmaterial3(); // ????3????
					String partthickness3 = wpb.getPartthickness3(); // ????3????
					String layersnum = wpb.getLayersnum(); // ??????
					String gagi = wpb.getGagi(); // GA /GI
					String sheetstrength440 = wpb.getSheetstrength440(); // ????????(Mpa)440
					String sheetstrength590 = wpb.getSheetstrength590(); // ????????(Mpa)590
					String sheetstrength = wpb.getSheetstrength(); // ????????(Mpa)>590
					String sheetstrength12 = wpb.getSheetstrength12(); // ????????(Mpa)1.2G
					String basethickness = wpb.getBasethickness(); // ????????
					String CurrentSerie = ""; // ???? ???? (????)
					String RecomWeldForce = "";// ???? ??????(N)
					String CurrentSeriedfi = ""; // ???? ???? (????)

					// ????????????????????????????????????????????
					boolean partmaterialFlag1 = true;
					boolean partmaterialFlag2 = true;
					boolean partmaterialFlag3 = true;
					String gagi1 = wpb.getGagi1();								
					String gagi2 = wpb.getGagi2();								
					String gagi3 = wpb.getGagi3();

					// ??????????????????GA/GI????
					if (MaterialMap != null) {
						for (Map.Entry<String, List<String>> entry : MaterialMap.entrySet()) {
							String MaterialNo = entry.getKey();
							List<String> infolist = entry.getValue();
							if(!"GA".equals(gagi1) && !"GI".equals(gagi1))
							{
								if (Util.getIsEqueal(partmaterial1, MaterialNo)) {
									if ("??".equalsIgnoreCase(infolist.get(1))) {
										partmaterialFlag1 = false;
									}
								}
							}
							if(!"GA".equals(gagi2) && !"GI".equals(gagi2))
							{
								if (Util.getIsEqueal(partmaterial2, MaterialNo)) {
									if ("??".equalsIgnoreCase(infolist.get(1))) {
										partmaterialFlag2 = false;
									}
								}
							}
							if(!"GA".equals(gagi3) && !"GI".equals(gagi3))
							{
								if (Util.getIsEqueal(partmaterial3, MaterialNo)) {
									if ("??".equalsIgnoreCase(infolist.get(1))) {
										partmaterialFlag3 = false;
									}
								}
							}																				
						}
					}
					if (partmaterialFlag1 && partmaterialFlag2 && partmaterialFlag3) {
						if (hjmap.containsKey(weldno)) {
							String[] curenre = hjmap.get(weldno);
							CurrentSerie = curenre[1];
							RecomWeldForce = curenre[4];
							CurrentSeriedfi = curenre[2];
						}
					}

					boolean flag = false;
					// ??????1.2g????????????????????????????
					if (sheetstrength12.equals("1.2g")) {
						flag = true;
					}
					if (flag) {
						basethickness = getMinnum(partthickness1, partthickness2, partthickness3);
					}
					setStringCellAndStyle(sheet, pageNo, 11 + j, 2, style, Cell.CELL_TYPE_STRING);
					setStringCellAndStyle(sheet, dot, 11 + j, 6, style8, Cell.CELL_TYPE_STRING);
					setStringCellAndStyle(sheet, importance, 11 + j, 4, style, Cell.CELL_TYPE_STRING);
					if (updateflag && rswsflist != null) {
						if (!rswsflist.contains(weldno)) {
							XSSFCellStyle newstyle = getXSSFStyle(book, sheet, 11 + j, 8, 2, -1);
							setStringCellAndStyle2(sheet, weldno, 11 + j, 8, newstyle, Cell.CELL_TYPE_STRING);
						} else {
							XSSFCellStyle newstyle = getXSSFStyle(book, sheet, 11 + j, 8, 12, -1);
							setStringCellAndStyle2(sheet, weldno, 11 + j, 8, newstyle, Cell.CELL_TYPE_STRING);
						}
					} else {
						XSSFCellStyle newstyle = getXSSFStyle(book, sheet, 11 + j, 8, 12, -1);
						setStringCellAndStyle2(sheet, weldno, 11 + j, 8, newstyle, Cell.CELL_TYPE_STRING);
					}
					setStringCellAndStyle(sheet, boardnumber1, 11 + j, 13, style, Cell.CELL_TYPE_STRING);
					setStringCellAndStyle(sheet, boardname1, 11 + j, 16, style, Cell.CELL_TYPE_STRING);
					if (getIscontains1180(partmaterial1)) {
						XSSFCellStyle newstyle = getXSSFStyleByrgb(book, sheet, 11 + j, 29, -1, new XSSFColor(new java.awt.Color(255,199,206)));
						setStringCellAndStyle2(sheet, partmaterial1, 11 + j, 29, newstyle, Cell.CELL_TYPE_STRING);
					} else {
						XSSFCellStyle newstyle = getXSSFStyle(book, sheet, 11 + j, 29, -1, IndexedColors.WHITE.getIndex());
						setStringCellAndStyle2(sheet, partmaterial1, 11 + j, 29, newstyle, Cell.CELL_TYPE_STRING);
					}
//					if (!partmaterialFlag1) {
//						setStringCellAndStyle2(sheet, partmaterial1, 11 + j, 29, stylepink, Cell.CELL_TYPE_STRING);
//					} else {
//						setStringCellAndStyle2(sheet, partmaterial1, 11 + j, 29, style, Cell.CELL_TYPE_STRING);
//					}
					setStringCellAndStyle(sheet, partthickness1, 11 + j, 36, style, 11);
					setStringCellAndStyle(sheet, boardnumber2, 11 + j, 39, style, Cell.CELL_TYPE_STRING);
					setStringCellAndStyle(sheet, boardname2, 11 + j, 42, style, Cell.CELL_TYPE_STRING);
					if (getIscontains1180(partmaterial2)) {
						XSSFCellStyle newstyle = getXSSFStyleByrgb(book, sheet, 11 + j, 55, -1, new XSSFColor(new java.awt.Color(255,199,206)));
						setStringCellAndStyle2(sheet, partmaterial2, 11 + j, 55, newstyle, Cell.CELL_TYPE_STRING);
					} else {
						XSSFCellStyle newstyle = getXSSFStyle(book, sheet, 11 + j, 55, -1, IndexedColors.WHITE.getIndex());
						setStringCellAndStyle2(sheet, partmaterial2, 11 + j, 55, newstyle, Cell.CELL_TYPE_STRING);
					}
//					if (!partmaterialFlag2) {
//						setStringCellAndStyle2(sheet, partmaterial2, 11 + j, 55, stylepink, Cell.CELL_TYPE_STRING);
//					} else {
//						setStringCellAndStyle2(sheet, partmaterial2, 11 + j, 55, style, Cell.CELL_TYPE_STRING);
//					}
					setStringCellAndStyle(sheet, partthickness2, 11 + j, 62, style, 11);
					setStringCellAndStyle(sheet, boardnumber3, 11 + j, 65, style, Cell.CELL_TYPE_STRING);
					setStringCellAndStyle(sheet, boardname3, 11 + j, 68, style, Cell.CELL_TYPE_STRING);
					if (getIscontains1180(partmaterial3)) {
						XSSFCellStyle newstyle = getXSSFStyleByrgb(book, sheet, 11 + j, 81, -1, new XSSFColor(new java.awt.Color(255,199,206)));
						setStringCellAndStyle2(sheet, partmaterial3, 11 + j, 81, newstyle, Cell.CELL_TYPE_STRING);
					} else {
						XSSFCellStyle newstyle = getXSSFStyle(book, sheet, 11 + j, 81, -1, IndexedColors.WHITE.getIndex());
						setStringCellAndStyle2(sheet, partmaterial3, 11 + j, 81, newstyle, Cell.CELL_TYPE_STRING);
					}
//					if (!partmaterialFlag3) {
//						setStringCellAndStyle2(sheet, partmaterial3, 11 + j, 81, stylepink, Cell.CELL_TYPE_STRING);
//					} else {
//						setStringCellAndStyle2(sheet, partmaterial3, 11 + j, 81, style, Cell.CELL_TYPE_STRING);
//					}
					setStringCellAndStyle(sheet, partthickness3, 11 + j, 87, style, 11);
					setStringCellAndStyle(sheet, layersnum, 11 + j, 90, style, 10);
					if (gagi != null && !gagi.isEmpty()) {
						setStringCellAndStyle(sheet, gagi, 11 + j, 92, style, Cell.CELL_TYPE_STRING);
					} else {
						setStringCellAndStyle(sheet, "-", 11 + j, 92, style, Cell.CELL_TYPE_STRING);
					}
					setStringCellAndStyle(sheet, sheetstrength440, 11 + j, 94, style, 10);
					setStringCellAndStyle(sheet, sheetstrength590, 11 + j, 96, style, 10);
					setStringCellAndStyle(sheet, sheetstrength, 11 + j, 98, style, 10);
					if (flag) {
						setStringCellAndStyle(sheet, "??", 11 + j, 100, style, Cell.CELL_TYPE_STRING);
						if (getColorDistinction(layersnum, partmaterial1, partmaterial2, partmaterial3, partthickness1,
								partthickness2, partthickness3)) {
							XSSFCellStyle newstyle = getXSSFStyle(book, sheet, 11 + j, 102, 1, IndexedColors.SKY_BLUE.getIndex());
							setStringCellAndStyle2(sheet, basethickness, 11 + j, 102, newstyle, 11);
						} else {
							XSSFCellStyle newstyle = getXSSFStyle(book, sheet, 11 + j, 102, 1, IndexedColors.VIOLET.getIndex());
							setStringCellAndStyle2(sheet, basethickness, 11 + j, 102, newstyle, 11);
						}
						// ????????????????
						setStringCellAndStyle(sheet, "", 11 + j, 105, style, Cell.CELL_TYPE_STRING);
						setStringCellAndStyle(sheet, "", 11 + j, 108, style, Cell.CELL_TYPE_STRING);
						setStringCellAndStyle(sheet, "", 11 + j, 111, style, Cell.CELL_TYPE_STRING);
					} else {
						XSSFCellStyle newstyle = getXSSFStyle(book, sheet, 11 + j, 102, 12, IndexedColors.WHITE.getIndex());
						setStringCellAndStyle(sheet, basethickness, 11 + j, 102, newstyle, 11);
						setStringCellAndStyle(sheet, "-", 11 + j, 100, style, Cell.CELL_TYPE_STRING);
						setStringCellAndStyle(sheet, CurrentSerie, 11 + j, 105, style, Cell.CELL_TYPE_STRING);
						setStringCellAndStyle(sheet, RecomWeldForce, 11 + j, 108, style, Cell.CELL_TYPE_STRING);
						setStringCellAndStyle(sheet, CurrentSeriedfi, 11 + j, 111, style, Cell.CELL_TYPE_STRING);
					}
				}
			} else {
				for (int j = 0; j + 36 * shnum < 36 + 36 * shnum; j++) {
//					WeldPointBoardInformation wpb = hdinfo.get(j + 36 * shnum);
					WeldPointBoardInformation wpb = new WeldPointBoardInformation();
					if (j + 36 * shnum > hdinfo.size() - 1) {
						wpb = symmetryhdinfo.get(j + 36 * shnum - hdinfo.size());
					} else {
						wpb = hdinfo.get(j + 36 * shnum);
					}
					String weldno = wpb.getWeldno(); // ????????
					String pageNo = "";// ????
					String dot = "";// ??????

					// ??????????????????????????
					if (notelist.containsKey(weldno)) {
						String[] pd = notelist.get(weldno);
						pageNo = pd[0];
						dot = pd[1];
					}
					String importance = wpb.getImportance(); // ??????
					// ??????????????????????????????
					if (importance != null && !importance.isEmpty()) {
						if (!Import.contains(importance)) {
							Import.add(importance);
						}
					}
					String boardnumber1 = wpb.getBoardnumber1(); // ????1????
					String boardname1 = wpb.getBoardname1(); // ????1????
					String partmaterial1 = wpb.getPartmaterial1(); // ????1????
					String partthickness1 = wpb.getPartthickness1(); // ????1????
					String boardnumber2 = wpb.getBoardnumber2(); // ????2????
					String boardname2 = wpb.getBoardname2(); // ????2????
					String partmaterial2 = wpb.getPartmaterial2(); // ????2????
					String partthickness2 = wpb.getPartthickness2(); // ????2????
					String boardnumber3 = wpb.getBoardnumber3(); // ????3????
					String boardname3 = wpb.getBoardname3(); // ????3????
					String partmaterial3 = wpb.getPartmaterial3(); // ????3????
					String partthickness3 = wpb.getPartthickness3(); // ????3????
					String layersnum = wpb.getLayersnum(); // ??????
					String gagi = wpb.getGagi(); // GA /GI
					String sheetstrength440 = wpb.getSheetstrength440(); // ????????(Mpa)440
					String sheetstrength590 = wpb.getSheetstrength590(); // ????????(Mpa)590
					String sheetstrength = wpb.getSheetstrength(); // ????????(Mpa)>590
					String sheetstrength12 = wpb.getSheetstrength12(); // ????????(Mpa)1.2G
					String basethickness = wpb.getBasethickness(); // ????????

					String CurrentSerie = ""; // ???? ???? (????)
					String RecomWeldForce = "";// ???? ??????(N)
					String CurrentSeriedfi = ""; // ???? ???? (????)

					// ????????????????????????????????????????????
					boolean partmaterialFlag1 = true;
					boolean partmaterialFlag2 = true;
					boolean partmaterialFlag3 = true;
					String gagi1 = wpb.getGagi1();								
					String gagi2 = wpb.getGagi2();								
					String gagi3 = wpb.getGagi3();

					// ??????????????????GA/GI????
					if (MaterialMap != null) {
						for (Map.Entry<String, List<String>> entry : MaterialMap.entrySet()) {
							String MaterialNo = entry.getKey();
							List<String> infolist = entry.getValue();
							if(!"GA".equals(gagi1) && !"GI".equals(gagi1))
							{
								if (Util.getIsEqueal(partmaterial1, MaterialNo)) {
									if ("??".equalsIgnoreCase(infolist.get(1))) {
										partmaterialFlag1 = false;
									}
								}
							}
							if(!"GA".equals(gagi2) && !"GI".equals(gagi2))
							{
								if (Util.getIsEqueal(partmaterial2, MaterialNo)) {
									if ("??".equalsIgnoreCase(infolist.get(1))) {
										partmaterialFlag2 = false;
									}
								}
							}
							if(!"GA".equals(gagi3) && !"GI".equals(gagi3))
							{
								if (Util.getIsEqueal(partmaterial3, MaterialNo)) {
									if ("??".equalsIgnoreCase(infolist.get(1))) {
										partmaterialFlag3 = false;
									}
								}
							}																				
						}
					}
					if (partmaterialFlag1 && partmaterialFlag2 && partmaterialFlag3) {
						if (hjmap.containsKey(weldno)) {
							String[] curenre = hjmap.get(weldno);
							CurrentSerie = curenre[1];
							RecomWeldForce = curenre[4];
							CurrentSeriedfi = curenre[2];
						}
					}
					boolean flag = false;
					// ??????1.2g????????????????????????????
					if (sheetstrength12.equals("1.2g")) {
						flag = true;
					}
					if (flag) {
						basethickness = getMinnum(partthickness1, partthickness2, partthickness3);
					}
					setStringCellAndStyle(sheet, pageNo, 11 + j, 2, style, Cell.CELL_TYPE_STRING);
					setStringCellAndStyle(sheet, dot, 11 + j, 6, style8, Cell.CELL_TYPE_STRING);
					setStringCellAndStyle(sheet, importance, 11 + j, 4, style, Cell.CELL_TYPE_STRING);
					if (updateflag && rswsflist != null) {
						if (!rswsflist.contains(weldno)) {
							XSSFCellStyle newstyle = getXSSFStyle(book, sheet, 11 + j, 8, 2, -1);
							setStringCellAndStyle2(sheet, weldno, 11 + j, 8, newstyle, Cell.CELL_TYPE_STRING);
						} else {
							XSSFCellStyle newstyle = getXSSFStyle(book, sheet, 11 + j, 8, 12, -1);
							setStringCellAndStyle2(sheet, weldno, 11 + j, 8, newstyle, Cell.CELL_TYPE_STRING);
						}
					} else {
						XSSFCellStyle newstyle = getXSSFStyle(book, sheet, 11 + j, 8, 12, -1);
						setStringCellAndStyle2(sheet, weldno, 11 + j, 8, newstyle, Cell.CELL_TYPE_STRING);
					}
					setStringCellAndStyle(sheet, boardnumber1, 11 + j, 13, style, Cell.CELL_TYPE_STRING);
					setStringCellAndStyle(sheet, boardname1, 11 + j, 16, style, Cell.CELL_TYPE_STRING);
					if (getIscontains1180(partmaterial1)) {
						XSSFCellStyle newstyle = getXSSFStyleByrgb(book, sheet, 11 + j, 29, -1, new XSSFColor(new java.awt.Color(255,199,206)));
						setStringCellAndStyle2(sheet, partmaterial1, 11 + j, 29, newstyle, Cell.CELL_TYPE_STRING);
					} else {
						XSSFCellStyle newstyle = getXSSFStyle(book, sheet, 11 + j, 29, -1, IndexedColors.WHITE.getIndex());
						setStringCellAndStyle2(sheet, partmaterial1, 11 + j, 29, newstyle, Cell.CELL_TYPE_STRING);
					}
//					if (!partmaterialFlag1) {
//						setStringCellAndStyle2(sheet, partmaterial1, 11 + j, 29, stylepink, Cell.CELL_TYPE_STRING);
//					} else {
//						setStringCellAndStyle2(sheet, partmaterial1, 11 + j, 29, style, Cell.CELL_TYPE_STRING);
//					}
					setStringCellAndStyle(sheet, partthickness1, 11 + j, 36, style, 11);
					setStringCellAndStyle(sheet, boardnumber2, 11 + j, 39, style, Cell.CELL_TYPE_STRING);
					setStringCellAndStyle(sheet, boardname2, 11 + j, 42, style, Cell.CELL_TYPE_STRING);
					if (getIscontains1180(partmaterial2)) {
						XSSFCellStyle newstyle = getXSSFStyleByrgb(book, sheet, 11 + j, 55, -1, new XSSFColor(new java.awt.Color(255,199,206)));
						setStringCellAndStyle2(sheet, partmaterial2, 11 + j, 55, newstyle, Cell.CELL_TYPE_STRING);
					} else {
						XSSFCellStyle newstyle = getXSSFStyle(book, sheet, 11 + j, 55, -1, IndexedColors.WHITE.getIndex());
						setStringCellAndStyle2(sheet, partmaterial2, 11 + j, 55, newstyle, Cell.CELL_TYPE_STRING);
					}
//					if (!partmaterialFlag2) {
//						setStringCellAndStyle2(sheet, partmaterial2, 11 + j, 55, stylepink, Cell.CELL_TYPE_STRING);
//					} else {
//						setStringCellAndStyle2(sheet, partmaterial2, 11 + j, 55, style, Cell.CELL_TYPE_STRING);
//					}
					setStringCellAndStyle(sheet, partthickness2, 11 + j, 62, style, 11);
					setStringCellAndStyle(sheet, boardnumber3, 11 + j, 65, style, Cell.CELL_TYPE_STRING);
					setStringCellAndStyle(sheet, boardname3, 11 + j, 68, style, Cell.CELL_TYPE_STRING);
					if (getIscontains1180(partmaterial3)) {
						XSSFCellStyle newstyle = getXSSFStyleByrgb(book, sheet, 11 + j, 81, -1, new XSSFColor(new java.awt.Color(255,199,206)));
						setStringCellAndStyle2(sheet, partmaterial3, 11 + j, 81, newstyle, Cell.CELL_TYPE_STRING);
					} else {
						XSSFCellStyle newstyle = getXSSFStyle(book, sheet, 11 + j, 81, -1, IndexedColors.WHITE.getIndex());
						setStringCellAndStyle2(sheet, partmaterial3, 11 + j, 81, newstyle, Cell.CELL_TYPE_STRING);
					}
//					if (!partmaterialFlag3) {
//						setStringCellAndStyle2(sheet, partmaterial3, 11 + j, 81, stylepink, Cell.CELL_TYPE_STRING);
//					} else {
//						setStringCellAndStyle2(sheet, partmaterial3, 11 + j, 81, style, Cell.CELL_TYPE_STRING);
//					}
					setStringCellAndStyle(sheet, partthickness3, 11 + j, 87, style, 11);
					setStringCellAndStyle(sheet, layersnum, 11 + j, 90, style, 10);
					if (gagi != null && !gagi.isEmpty()) {
						setStringCellAndStyle(sheet, gagi, 11 + j, 92, style, Cell.CELL_TYPE_STRING);
					} else {
						setStringCellAndStyle(sheet, "-", 11 + j, 92, style, Cell.CELL_TYPE_STRING);
					}
					setStringCellAndStyle(sheet, sheetstrength440, 11 + j, 94, style, 10);
					setStringCellAndStyle(sheet, sheetstrength590, 11 + j, 96, style, 10);
					setStringCellAndStyle(sheet, sheetstrength, 11 + j, 98, style, 10);
					setStringCellAndStyle(sheet, sheetstrength, 11 + j, 98, style, 10);
					if (flag) {
						setStringCellAndStyle(sheet, "??", 11 + j, 100, style, Cell.CELL_TYPE_STRING);
						if (getColorDistinction(layersnum, partmaterial1, partmaterial2, partmaterial3, partthickness1,
								partthickness2, partthickness3)) {
							XSSFCellStyle newstyle = getXSSFStyle(book, sheet, 11 + j, 102, 1, IndexedColors.SKY_BLUE.getIndex());
							setStringCellAndStyle2(sheet, basethickness, 11 + j, 102, newstyle, 11);
						} else {
							XSSFCellStyle newstyle = getXSSFStyle(book, sheet, 11 + j, 102, 1, IndexedColors.VIOLET.getIndex());
							setStringCellAndStyle2(sheet, basethickness, 11 + j, 102, newstyle, 11);
						}
						// ????????????????
						setStringCellAndStyle(sheet, "", 11 + j, 105, style, Cell.CELL_TYPE_STRING);
						setStringCellAndStyle(sheet, "", 11 + j, 108, style, Cell.CELL_TYPE_STRING);
						setStringCellAndStyle(sheet, "", 11 + j, 111, style, Cell.CELL_TYPE_STRING);
					} else {
						XSSFCellStyle newstyle = getXSSFStyle(book, sheet, 11 + j, 102, 12, IndexedColors.WHITE.getIndex());
						setStringCellAndStyle(sheet, basethickness, 11 + j, 102, newstyle, 11);
						setStringCellAndStyle(sheet, "-", 11 + j, 100, style, Cell.CELL_TYPE_STRING);
						setStringCellAndStyle(sheet, CurrentSerie, 11 + j, 105, style, Cell.CELL_TYPE_STRING);
						setStringCellAndStyle(sheet, RecomWeldForce, 11 + j, 108, style, Cell.CELL_TYPE_STRING);
						setStringCellAndStyle(sheet, CurrentSeriedfi, 11 + j, 111, style, Cell.CELL_TYPE_STRING);
					}
				}
			}

			shnum++;
		}
	}

	private Comparator getComParatorByfirstpart() {
		// TODO Auto-generated method stub
		Comparator comparator = new Comparator() {

			public int compare(Object obj, Object obj1) {
				// System.setProperty("java.util.Arrays.useLegacyMergeSort", "true");
				WeldPointBoardInformation comp1 = (WeldPointBoardInformation) obj;
				WeldPointBoardInformation comp2 = (WeldPointBoardInformation) obj1;

				String d1 = "";
				String d2 = "";
				if (obj != null && comp1.getBoardnumber1() != null && !comp1.getBoardnumber1().isEmpty()) {
					d1 = comp1.getBoardnumber1().toString();
				}
				if (obj1 != null && comp2.getBoardnumber1() != null && !comp2.getBoardnumber1().isEmpty()) {
					d2 = comp2.getBoardnumber1().toString();
					;
				}
				return d1.compareTo(d2);
			}
		};

		return comparator;
	}

	private Comparator getComParatorBySecondpart() {
		// TODO Auto-generated method stub
		Comparator comparator = new Comparator() {

			public int compare(Object obj, Object obj1) {
				// System.setProperty("java.util.Arrays.useLegacyMergeSort", "true");
				WeldPointBoardInformation comp1 = (WeldPointBoardInformation) obj;
				WeldPointBoardInformation comp2 = (WeldPointBoardInformation) obj1;

				String d1 = "";
				String d2 = "";
				if (obj != null && comp1.getBoardnumber2() != null && !comp1.getBoardnumber2().isEmpty()) {
					d1 = comp1.getBoardnumber2().toString();
				}
				if (obj1 != null && comp2.getBoardnumber2() != null && !comp2.getBoardnumber2().isEmpty()) {
					d2 = comp2.getBoardnumber2().toString();
					;
				}
				return d1.compareTo(d2);
			}
		};

		return comparator;
	}

	private Comparator getComParatorByThistypart() {
		// TODO Auto-generated method stub
		Comparator comparator = new Comparator() {

			public int compare(Object obj, Object obj1) {
				// System.setProperty("java.util.Arrays.useLegacyMergeSort", "true");
				WeldPointBoardInformation comp1 = (WeldPointBoardInformation) obj;
				WeldPointBoardInformation comp2 = (WeldPointBoardInformation) obj1;

				String d1 = "";
				String d2 = "";
				if (obj != null && comp1.getBoardnumber3() != null && !comp1.getBoardnumber3().isEmpty()) {
					d1 = comp1.getBoardnumber3().toString();
				}
				if (obj1 != null && comp2.getBoardnumber3() != null && !comp2.getBoardnumber3().isEmpty()) {
					d2 = comp2.getBoardnumber3().toString();
					;
				}
				return d1.compareTo(d2);
			}
		};

		return comparator;
	}

	/*
	 * RSW????????????
	 */
	private void RSWpneumaticinformationProcessing(XSSFWorkbook book, TCComponentBOMLine bl, TCComponentBOMLine gwbl,
			String name, Map<String, TCComponentBOMLine> symmetrymap) throws TCException {
		// TODO Auto-generated method stub
		// ????RSW????sheet
		int sheetnum = 0;
		sheetnum = book.getNumberOfSheets();
		int sheetAtIndex = -1; // RSW????????????
		for (int i = 0; i < sheetnum; i++) {
			String sheetname = book.getSheetName(i);
			if (sheetname.contains(name)) {
				sheetAtIndex = i;
				break;
			}
		}
		if (sheetAtIndex == -1) {
			return;
		}
		// ??????????????????????
		ArrayList gunlist = new ArrayList();
		ArrayList hdlist = new ArrayList();
		// ArrayList robotlist = new ArrayList();
		String guntypename = Util.getObjectDisplayName(session, "B8_BIWGun");
		String[] propertys = new String[] { "bl_item_object_type", "bl_item_object_type" };
		String[] values = new String[] { guntypename, guntypename };
		String weldtypename = Util.getObjectDisplayName(session, "WeldPoint");
		String[] propertys2 = new String[] { "bl_item_object_type", "bl_item_object_type" };
		String[] values2 = new String[] { weldtypename, weldtypename };

//		String[] propertys3 = new String[] { "bl_item_object_type", "bl_item_object_type" };
//		String[] values3 = new String[] { "??????", "BIW Robot" };

		gunlist = Util.searchBOMLine(bl, "OR", propertys, "==", values);
		hdlist = Util.searchBOMLine(bl, "OR", propertys2, "==", values2);
		// robotlist = Util.searchBOMLine(bl, "OR", propertys3, "==", values3);

		// ??????????????SOP????????????SOP??????????????????
		boolean sopflag = getIsSOPAfter();
		String Discretename = Util.getProperty(bl, "bl_rev_object_name");

		// ????????????????????????
		String stationname = Util.getProperty(gwbl, "bl_rev_object_name");// ????
		String robotname = "";// ??????
		String Guncode = "";// ????????
		String ElectrodeVol = "";// ???????? bl_B8_BIWGunRevision_b8_ElectrodeVol
		TCComponentItemRevision blrev = null;
		if (gunlist.size() > 0) {
			TCComponentBOMLine gunbl = (TCComponentBOMLine) gunlist.get(0);
			blrev = gunbl.getItemRevision();
			Guncode = Util.getProperty(blrev, "b8_Model");
			ElectrodeVol = Util.getProperty(blrev, "b8_ElectrodeVol");
		}
		robotname = Util.getProperty(bl, "bl_rev_object_name");
//		if (robotlist.size() > 0) {
//			TCComponentBOMLine robotbl = (TCComponentBOMLine) robotlist.get(0);
//			robotname = Util.getProperty(robotbl, "bl_rev_object_name");
//		}
		// ??????????????????????????????????
		String[] nameArr = Discretename.split("\\\\");
		if (nameArr.length > 1) {
			Guncode = nameArr[1];
		} else {
			Guncode = "";
		}
		robotname = nameArr[0];

		// ??????????????????????????????????
		List<WeldPointBoardInformation> hdinfo = new ArrayList<WeldPointBoardInformation>();// ????????????

		hdinfo = getBoardInformation(baseinfolist, hdlist);

		// ????????????
		Map<String, String[]> hjmap = getHanJieParater(hdlist);

		if (hdinfo == null || hdinfo.size() < 1) {
			System.out.println("????????????????????????????");
			return;
		}

		// ????????????????????
		Comparator comparator = getComParatorByfirstpart();
		Collections.sort(hdinfo, comparator);

		// ??????????????????????
		Comparator comparator1 = getComParatorBySecondpart();
		Collections.sort(hdinfo, comparator1);

		// ????????????????????
		Comparator comparator11 = getComParatorByThistypart();
		Collections.sort(hdinfo, comparator11);

		// ??????????????????????3????????????2????
		Comparator comparator2 = getComParatorBylayersnum();
		Collections.sort(hdinfo, comparator2);

		// ??????????????????????????????????????????????????????????????????????????????????sheet??????
//		String Discretename = Util.getProperty(bl, "bl_rev_object_name");
		// ??????????????????????????????????
		List<WeldPointBoardInformation> symmetryhdinfo = new ArrayList<WeldPointBoardInformation>();// ????????????
		ArrayList symmetryhdlist = new ArrayList();
		if (symmetrymap.containsKey(Discretename)) {
			TCComponentBOMLine symmetrybl = symmetrymap.get(Discretename);
			symmetryhdlist = Util.searchBOMLine(symmetrybl, "OR", propertys2, "==", values2);
			symmetryhdinfo = getBoardInformation(baseinfolist, symmetryhdlist);

			// ????????????????????
			Comparator comparators1 = getComParatorByfirstpart();
			Collections.sort(symmetryhdinfo, comparators1);
			// ??????????????????????
			Comparator comparators2 = getComParatorBySecondpart();
			Collections.sort(symmetryhdinfo, comparators2);

			// ????????????????????
			Comparator comparators3 = getComParatorByThistypart();
			Collections.sort(symmetryhdinfo, comparators3);

			// ??????????????????????3????????????2????
			Comparator comparators4 = getComParatorBylayersnum();
			Collections.sort(symmetryhdinfo, comparators4);

			// ??????????????????
			getSymmetryHanJieParater(symmetryhdlist, hjmap);
		}

		// ????????????????????????,??36????????????
		int hdsum = hdlist.size() + symmetryhdlist.size();
		int page = hdsum / 36 + 1;
		// ??????????????????????sheet????????????????
		if (hdsum % 36 == 0) {
			if (page > 1) {
				page = page - 1;
			}
		}
		// ????page????1????????????sheet??
		int index = sheetAtIndex + 1;

		// ????????????
		Font font = book.createFont();
		font.setColor((short) 12);// ????????
		font.setFontHeightInPoints((short) 9);
		XSSFCellStyle style = book.createCellStyle();
		style.setBorderBottom(XSSFCellStyle.BORDER_THIN); // ??????
		style.setBorderLeft(XSSFCellStyle.BORDER_THIN);// ??????
		style.setBorderTop(XSSFCellStyle.BORDER_THIN);// ??????
		style.setBorderRight(XSSFCellStyle.BORDER_THIN);// ??????
		style.setVerticalAlignment(XSSFCellStyle.VERTICAL_CENTER);
		style.setAlignment(XSSFCellStyle.ALIGN_CENTER);
		style.setFont(font);

		Font font2 = book.createFont();
		font2.setColor((short) 12);// ????????
		font2.setFontHeightInPoints((short) 18);
		font2.setBoldweight(Font.BOLDWEIGHT_BOLD); // ????????
		XSSFCellStyle style2 = book.createCellStyle();
		style2.setBorderBottom(CellStyle.BORDER_MEDIUM); // ????????
		style2.setBorderLeft(CellStyle.BORDER_MEDIUM); // ????????
		style2.setBorderRight(CellStyle.BORDER_MEDIUM); // ????????
		style2.setBorderTop(CellStyle.BORDER_MEDIUM); // ????????
		style2.setVerticalAlignment(XSSFCellStyle.VERTICAL_CENTER);
		style2.setAlignment(XSSFCellStyle.ALIGN_CENTER);
		style2.setFont(font2);

		XSSFCellStyle style3 = book.createCellStyle();
		style3.setBorderBottom(XSSFCellStyle.BORDER_THIN); // ??????
		// style3.setBorderLeft(XSSFCellStyle.BORDER_THIN);// ??????
		style3.setBorderTop(XSSFCellStyle.BORDER_THIN);// ??????
		style3.setBorderRight(XSSFCellStyle.BORDER_THIN);// ??????
		style3.setVerticalAlignment(XSSFCellStyle.VERTICAL_CENTER);
		style3.setAlignment(XSSFCellStyle.ALIGN_CENTER);
		style3.setFont(font);

		XSSFCellStyle style4 = book.createCellStyle();
		style4.setBorderBottom(XSSFCellStyle.BORDER_THIN); // ??????
		style4.setBorderLeft(XSSFCellStyle.BORDER_THIN);// ??????
		style4.setBorderTop(XSSFCellStyle.BORDER_THIN);// ??????
		// style4.setBorderRight(XSSFCellStyle.BORDER_THIN);// ??????
		style4.setVerticalAlignment(XSSFCellStyle.VERTICAL_CENTER);
		style4.setAlignment(XSSFCellStyle.ALIGN_CENTER);
		style4.setFont(font);

		Font font3 = book.createFont();
		font3.setColor((short) 2);// ????????
		font3.setFontHeightInPoints((short) 10);
		XSSFCellStyle style33 = book.createCellStyle();
		style33.setBorderBottom(XSSFCellStyle.BORDER_THIN); // ??????
		// style3.setBorderLeft(XSSFCellStyle.BORDER_THIN);// ??????
		style33.setBorderTop(XSSFCellStyle.BORDER_THIN);// ??????
		style33.setBorderRight(XSSFCellStyle.BORDER_THIN);// ??????
		style33.setVerticalAlignment(XSSFCellStyle.VERTICAL_CENTER);
		style33.setAlignment(XSSFCellStyle.ALIGN_CENTER);
		style33.setFont(font3);

		// ??????????
		Font fontpink = book.createFont();
		fontpink.setColor((short) 12);// ????????
		fontpink.setFontName("MS PGothic");
		fontpink.setFontHeightInPoints((short) 9);

		XSSFCellStyle stylepink = book.createCellStyle();
		stylepink.setFillForegroundColor(IndexedColors.ROSE.getIndex());
		stylepink.setFillPattern(CellStyle.SOLID_FOREGROUND);
		stylepink.setBorderBottom(XSSFCellStyle.BORDER_THIN); // ??????
		stylepink.setBorderLeft(XSSFCellStyle.BORDER_THIN);// ??????
		stylepink.setBorderTop(XSSFCellStyle.BORDER_THIN);// ??????
		stylepink.setBorderRight(XSSFCellStyle.BORDER_THIN);// ??????
		stylepink.setVerticalAlignment(XSSFCellStyle.VERTICAL_CENTER);
		stylepink.setAlignment(XSSFCellStyle.ALIGN_CENTER);
		stylepink.setFont(fontpink);

		if (updateflag) {
			int number = 0;
			for (int i = 0; i < sheetnum; i++) {
				String sheetname = book.getSheetName(i);
				if (sheetname.contains("RSW????")) {
					if (sheetAtIndex <= i) {
						number++;
					}
				}
			}
			index = index + page - 1;
			if (number < page) {
				if (page - number > 0) {
					for (int i = 0; i < page - number; i++) {
						XSSFSheet newsheet = book.cloneSheet(sheetAtIndex);
						book.setSheetOrder(newsheet.getSheetName(), index);
						index++;
					}
				}
			}
		} else {
			if (page > 1) {
				for (int i = 1; i < page; i++) {
					XSSFSheet newsheet = book.cloneSheet(sheetAtIndex);
					book.setSheetOrder(newsheet.getSheetName(), index);
					index++;
				}
			}
		}

		int shnum = 0;

		for (int i = sheetAtIndex; i < index; i++) {
			XSSFSheet sheet = book.getSheetAt(i);

			setStringCellAndStyle(sheet, stationname, 6, 19, style2, Cell.CELL_TYPE_STRING);// ????
			setStringCellAndStyle(sheet, robotname, 6, 30, style2, Cell.CELL_TYPE_STRING);// ??????
			setStringCellAndStyle(sheet, Guncode, 6, 47, style2, Cell.CELL_TYPE_STRING);// ????????
			setStringCellAndStyle(sheet, ElectrodeVol, 6, 64, style2, Cell.CELL_TYPE_STRING);// ????????
																								// bl_B8_BIWGunRevision_b8_ElectrodeVol

			if (i == index - 1) {
				for (int j = 0; j + 36 * shnum < hdinfo.size() + symmetryhdinfo.size(); j++) {
//					WeldPointBoardInformation wpb = hdinfo.get(j + 36 * shnum);
					WeldPointBoardInformation wpb = new WeldPointBoardInformation();
					if (j + 36 * shnum > hdinfo.size() - 1) {
						wpb = symmetryhdinfo.get(j + 36 * shnum - hdinfo.size());
					} else {
						wpb = hdinfo.get(j + 36 * shnum);
					}
					String weldno = wpb.getWeldno(); // ????????
					String pageNo = "";// ????
					String dot = "";// ??????

					// ??????????????????????????
					if (notelist.containsKey(weldno)) {
						String[] pd = notelist.get(weldno);
						pageNo = pd[0];
						dot = pd[1];
					}
					String importance = wpb.getImportance(); // ??????
					// ??????????????????????????????
					if (importance != null && !importance.isEmpty()) {
						if (!Import.contains(importance)) {
							Import.add(importance);
						}
					}
					String boardnumber1 = wpb.getBoardnumber1(); // ????1????
					String boardname1 = wpb.getBoardname1(); // ????1????
					String partmaterial1 = wpb.getPartmaterial1(); // ????1????
					String partthickness1 = wpb.getPartthickness1(); // ????1????
					String boardnumber2 = wpb.getBoardnumber2(); // ????2????
					String boardname2 = wpb.getBoardname2(); // ????2????
					String partmaterial2 = wpb.getPartmaterial2(); // ????2????
					String partthickness2 = wpb.getPartthickness2(); // ????2????
					String boardnumber3 = wpb.getBoardnumber3(); // ????3????
					String boardname3 = wpb.getBoardname3(); // ????3????
					String partmaterial3 = wpb.getPartmaterial3(); // ????3????
					String partthickness3 = wpb.getPartthickness3(); // ????3????
					String layersnum = wpb.getLayersnum(); // ??????
					String gagi = wpb.getGagi(); // GA /GI
					String sheetstrength440 = wpb.getSheetstrength440(); // ????????(Mpa)440
					String sheetstrength590 = wpb.getSheetstrength590(); // ????????(Mpa)590
					String sheetstrength = wpb.getSheetstrength(); // ????????(Mpa)>590
					String basethickness = wpb.getBasethickness(); // ????????
					String sheetstrength12 = wpb.getSheetstrength12(); // ????????(Mpa)1.2G

					String CurrentSerie = ""; // ???? ???? (????)
					String RecomWeldForce = "";// ???? ??????(N)

					// ????????????????????????????????????????????
					boolean partmaterialFlag1 = true;
					boolean partmaterialFlag2 = true;
					boolean partmaterialFlag3 = true;
					String gagi1 = wpb.getGagi1();								
					String gagi2 = wpb.getGagi2();								
					String gagi3 = wpb.getGagi3();

					// ??????????????????GA/GI????
					if (MaterialMap != null) {
						for (Map.Entry<String, List<String>> entry : MaterialMap.entrySet()) {
							String MaterialNo = entry.getKey();
							List<String> infolist = entry.getValue();
							if(!"GA".equals(gagi1) && !"GI".equals(gagi1))
							{
								if (Util.getIsEqueal(partmaterial1, MaterialNo)) {
									if ("??".equalsIgnoreCase(infolist.get(1))) {
										partmaterialFlag1 = false;
									}
								}
							}
							if(!"GA".equals(gagi2) && !"GI".equals(gagi2))
							{
								if (Util.getIsEqueal(partmaterial2, MaterialNo)) {
									if ("??".equalsIgnoreCase(infolist.get(1))) {
										partmaterialFlag2 = false;
									}
								}
							}
							if(!"GA".equals(gagi3) && !"GI".equals(gagi3))
							{
								if (Util.getIsEqueal(partmaterial3, MaterialNo)) {
									if ("??".equalsIgnoreCase(infolist.get(1))) {
										partmaterialFlag3 = false;
									}
								}
							}																				
						}
					}
					if (partmaterialFlag1 && partmaterialFlag2 && partmaterialFlag3) {
						if (hjmap.containsKey(weldno)) {
							String[] curenre = hjmap.get(weldno);
							CurrentSerie = curenre[1];
							RecomWeldForce = curenre[4];
						}
					}

					setStringCellAndStyle(sheet, pageNo, 11 + j, 2, style, Cell.CELL_TYPE_STRING);
					setStringCellAndStyle(sheet, dot, 11 + j, 6, style4, Cell.CELL_TYPE_STRING);
					setStringCellAndStyle(sheet, importance, 11 + j, 4, style, Cell.CELL_TYPE_STRING);
					if (updateflag && rswqdlist != null) {
						if (!rswqdlist.contains(weldno)) {
							XSSFCellStyle newstyle = getXSSFStyle(book, sheet, 11 + j, 8, 2, -1);
							setStringCellAndStyle2(sheet, weldno, 11 + j, 8, newstyle, Cell.CELL_TYPE_STRING);
						} else {
							XSSFCellStyle newstyle = getXSSFStyle(book, sheet, 11 + j, 8, 12, -1);
							setStringCellAndStyle2(sheet, weldno, 11 + j, 8, newstyle, Cell.CELL_TYPE_STRING);
						}
					} else {
						XSSFCellStyle newstyle = getXSSFStyle(book, sheet, 11 + j, 8, 12, -1);
						setStringCellAndStyle2(sheet, weldno, 11 + j, 8, newstyle, Cell.CELL_TYPE_STRING);
					}

					setStringCellAndStyle(sheet, boardnumber1, 11 + j, 13, style, Cell.CELL_TYPE_STRING);
					setStringCellAndStyle(sheet, boardname1, 11 + j, 16, style, Cell.CELL_TYPE_STRING);
//					if (!partmaterialFlag1) {
//						setStringCellAndStyle2(sheet, partmaterial1, 11 + j, 29, stylepink, Cell.CELL_TYPE_STRING);
//					} else {
//						setStringCellAndStyle2(sheet, partmaterial1, 11 + j, 29, style, Cell.CELL_TYPE_STRING);
//					}
					setStringCellAndStyle(sheet, partmaterial1, 11 + j, 29, style, Cell.CELL_TYPE_STRING);
					
					setStringCellAndStyle(sheet, partthickness1, 11 + j, 36, style, 11);
					setStringCellAndStyle(sheet, boardnumber2, 11 + j, 39, style, Cell.CELL_TYPE_STRING);
					setStringCellAndStyle(sheet, boardname2, 11 + j, 42, style, Cell.CELL_TYPE_STRING);
//					if (!partmaterialFlag2) {
//						setStringCellAndStyle2(sheet, partmaterial2, 11 + j, 55, stylepink, Cell.CELL_TYPE_STRING);
//					} else {
//						setStringCellAndStyle2(sheet, partmaterial2, 11 + j, 55, style, Cell.CELL_TYPE_STRING);
//					}
					setStringCellAndStyle(sheet, partmaterial2, 11 + j, 55, style, Cell.CELL_TYPE_STRING);
					
					setStringCellAndStyle(sheet, partthickness2, 11 + j, 62, style, 11);
					setStringCellAndStyle(sheet, boardnumber3, 11 + j, 65, style, Cell.CELL_TYPE_STRING);
					setStringCellAndStyle(sheet, boardname3, 11 + j, 68, style, Cell.CELL_TYPE_STRING);
//					if (!partmaterialFlag3) {
//						setStringCellAndStyle2(sheet, partmaterial3, 11 + j, 81, stylepink, Cell.CELL_TYPE_STRING);
//					} else {
//						setStringCellAndStyle2(sheet, partmaterial3, 11 + j, 81, style, Cell.CELL_TYPE_STRING);
//					}
					setStringCellAndStyle(sheet, partmaterial3, 11 + j, 81, style, Cell.CELL_TYPE_STRING);

					setStringCellAndStyle(sheet, partthickness3, 11 + j, 88, style, 11);
					setStringCellAndStyle(sheet, layersnum, 11 + j, 91, style, 10);
					if (gagi != null && !gagi.isEmpty()) {
						setStringCellAndStyle(sheet, gagi, 11 + j, 93, style, Cell.CELL_TYPE_STRING);
					} else {
						setStringCellAndStyle(sheet, "-", 11 + j, 93, style, Cell.CELL_TYPE_STRING);
					}

					setStringCellAndStyle(sheet, sheetstrength440, 11 + j, 95, style, 10);
					setStringCellAndStyle(sheet, sheetstrength590, 11 + j, 97, style, 10);
					setStringCellAndStyle(sheet, sheetstrength, 11 + j, 99, style, 10);
					setStringCellAndStyle(sheet, basethickness, 11 + j, 102, style, 11);
					// ??????1.2g????????????????????????????
					if (sheetstrength12.equals("1.2g")) {
						setStringCellAndStyle(sheet, "", 11 + j, 105, style, 10);
						setStringCellAndStyle(sheet, "", 11 + j, 108, style, 10);
						setStringCellAndStyle(sheet, "", 11 + j, 111, style, 10);
					} else {
						setStringCellAndStyle(sheet, CurrentSerie, 11 + j, 105, style, 10);
						setStringCellAndStyle(sheet, RecomWeldForce, 11 + j, 108, style, 10);
						setStringCellAndStyle(sheet, CurrentSerie, 11 + j, 111, style, 10);
					}
				}
			} else {
				for (int j = 0; j + 36 * shnum < 36 + 36 * shnum; j++) {
//					WeldPointBoardInformation wpb = hdinfo.get(j + 36 * shnum);
					WeldPointBoardInformation wpb = new WeldPointBoardInformation();
					if (j + 36 * shnum > hdinfo.size() - 1) {
						wpb = symmetryhdinfo.get(j + 36 * shnum - hdinfo.size());
					} else {
						wpb = hdinfo.get(j + 36 * shnum);
					}
					String weldno = wpb.getWeldno(); // ????????
					String pageNo = "";// ????
					String dot = "";// ??????

					// ??????????????????????????
					if (notelist.containsKey(weldno)) {
						String[] pd = notelist.get(weldno);
						pageNo = pd[0];
						dot = pd[1];
					}
					String importance = wpb.getImportance(); // ??????
					// ??????????????????????????????
					if (importance != null && !importance.isEmpty()) {
						if (!Import.contains(importance)) {
							Import.add(importance);
						}
					}
					String boardnumber1 = wpb.getBoardnumber1(); // ????1????
					String boardname1 = wpb.getBoardname1(); // ????1????
					String partmaterial1 = wpb.getPartmaterial1(); // ????1????
					String partthickness1 = wpb.getPartthickness1(); // ????1????
					String boardnumber2 = wpb.getBoardnumber2(); // ????2????
					String boardname2 = wpb.getBoardname2(); // ????2????
					String partmaterial2 = wpb.getPartmaterial2(); // ????2????
					String partthickness2 = wpb.getPartthickness2(); // ????2????
					String boardnumber3 = wpb.getBoardnumber3(); // ????3????
					String boardname3 = wpb.getBoardname3(); // ????3????
					String partmaterial3 = wpb.getPartmaterial3(); // ????3????
					String partthickness3 = wpb.getPartthickness3(); // ????3????
					String layersnum = wpb.getLayersnum(); // ??????
					String gagi = wpb.getGagi(); // GA /GI
					String sheetstrength440 = wpb.getSheetstrength440(); // ????????(Mpa)440
					String sheetstrength590 = wpb.getSheetstrength590(); // ????????(Mpa)590
					String sheetstrength = wpb.getSheetstrength(); // ????????(Mpa)>590
					String basethickness = wpb.getBasethickness(); // ????????
					String sheetstrength12 = wpb.getSheetstrength12(); // ????????(Mpa)1.2G

					String CurrentSerie = ""; // ???? ???? (????)
					String RecomWeldForce = "";// ???? ??????(N)

					// ????????????????????????????????????????????
					boolean partmaterialFlag1 = true;
					boolean partmaterialFlag2 = true;
					boolean partmaterialFlag3 = true;
					String gagi1 = wpb.getGagi1();								
					String gagi2 = wpb.getGagi2();								
					String gagi3 = wpb.getGagi3();

					// ??????????????????GA/GI????
					if (MaterialMap != null) {
						for (Map.Entry<String, List<String>> entry : MaterialMap.entrySet()) {
							String MaterialNo = entry.getKey();
							List<String> infolist = entry.getValue();
							if(!"GA".equals(gagi1) && !"GI".equals(gagi1))
							{
								if (Util.getIsEqueal(partmaterial1, MaterialNo)) {
									if ("??".equalsIgnoreCase(infolist.get(1))) {
										partmaterialFlag1 = false;
									}
								}
							}
							if(!"GA".equals(gagi2) && !"GI".equals(gagi2))
							{
								if (Util.getIsEqueal(partmaterial2, MaterialNo)) {
									if ("??".equalsIgnoreCase(infolist.get(1))) {
										partmaterialFlag2 = false;
									}
								}
							}
							if(!"GA".equals(gagi3) && !"GI".equals(gagi3))
							{
								if (Util.getIsEqueal(partmaterial3, MaterialNo)) {
									if ("??".equalsIgnoreCase(infolist.get(1))) {
										partmaterialFlag3 = false;
									}
								}
							}																				
						}
					}
					if (partmaterialFlag1 && partmaterialFlag2 && partmaterialFlag3) {
						if (hjmap.containsKey(weldno)) {
							String[] curenre = hjmap.get(weldno);
							CurrentSerie = curenre[1];
							RecomWeldForce = curenre[4];
						}
					}

					setStringCellAndStyle(sheet, pageNo, 11 + j, 2, style, Cell.CELL_TYPE_STRING);
					setStringCellAndStyle(sheet, dot, 11 + j, 6, style4, Cell.CELL_TYPE_STRING);
					setStringCellAndStyle(sheet, importance, 11 + j, 4, style, Cell.CELL_TYPE_STRING);
					if (updateflag && rswqdlist != null) {
						if (!rswqdlist.contains(weldno)) {
							XSSFCellStyle newstyle = getXSSFStyle(book, sheet, 11 + j, 8, 2, -1);
							setStringCellAndStyle2(sheet, weldno, 11 + j, 8, newstyle, Cell.CELL_TYPE_STRING);
						} else {
							XSSFCellStyle newstyle = getXSSFStyle(book, sheet, 11 + j, 8, 12, -1);
							setStringCellAndStyle2(sheet, weldno, 11 + j, 8, newstyle, Cell.CELL_TYPE_STRING);
						}
					} else {
						XSSFCellStyle newstyle = getXSSFStyle(book, sheet, 11 + j, 8, 12, -1);
						setStringCellAndStyle2(sheet, weldno, 11 + j, 8, newstyle, Cell.CELL_TYPE_STRING);
					}
					setStringCellAndStyle(sheet, boardnumber1, 11 + j, 13, style, Cell.CELL_TYPE_STRING);
					setStringCellAndStyle(sheet, boardname1, 11 + j, 16, style, Cell.CELL_TYPE_STRING);
//					if (!partmaterialFlag1) {
//						setStringCellAndStyle2(sheet, partmaterial1, 11 + j, 29, stylepink, Cell.CELL_TYPE_STRING);
//					} else {
//						setStringCellAndStyle2(sheet, partmaterial1, 11 + j, 29, style, Cell.CELL_TYPE_STRING);
//					}
					setStringCellAndStyle(sheet, partmaterial1, 11 + j, 29, style, Cell.CELL_TYPE_STRING);

					setStringCellAndStyle(sheet, partthickness1, 11 + j, 36, style, 11);
					setStringCellAndStyle(sheet, boardnumber2, 11 + j, 39, style, Cell.CELL_TYPE_STRING);
					setStringCellAndStyle(sheet, boardname2, 11 + j, 42, style, Cell.CELL_TYPE_STRING);
//					if (!partmaterialFlag2) {
//						setStringCellAndStyle2(sheet, partmaterial2, 11 + j, 55, stylepink, Cell.CELL_TYPE_STRING);
//					} else {
//						setStringCellAndStyle2(sheet, partmaterial2, 11 + j, 55, style, Cell.CELL_TYPE_STRING);
//					}
					setStringCellAndStyle(sheet, partmaterial2, 11 + j, 55, style, Cell.CELL_TYPE_STRING);

					setStringCellAndStyle(sheet, partthickness2, 11 + j, 62, style, 11);
					setStringCellAndStyle(sheet, boardnumber3, 11 + j, 65, style, Cell.CELL_TYPE_STRING);
					setStringCellAndStyle(sheet, boardname3, 11 + j, 68, style, Cell.CELL_TYPE_STRING);
//					if (!partmaterialFlag3) {
//						setStringCellAndStyle2(sheet, partmaterial3, 11 + j, 81, stylepink, Cell.CELL_TYPE_STRING);
//					} else {
//						setStringCellAndStyle2(sheet, partmaterial3, 11 + j, 81, style, Cell.CELL_TYPE_STRING);
//					}
					setStringCellAndStyle(sheet, partmaterial3, 11 + j, 81, style, Cell.CELL_TYPE_STRING);

					setStringCellAndStyle(sheet, partthickness3, 11 + j, 88, style, 11);
					setStringCellAndStyle(sheet, layersnum, 11 + j, 91, style, 10);
					if (gagi != null && !gagi.isEmpty()) {
						setStringCellAndStyle(sheet, gagi, 11 + j, 93, style, Cell.CELL_TYPE_STRING);
					} else {
						setStringCellAndStyle(sheet, "-", 11 + j, 93, style, Cell.CELL_TYPE_STRING);
					}
					setStringCellAndStyle(sheet, sheetstrength440, 11 + j, 95, style, 10);
					setStringCellAndStyle(sheet, sheetstrength590, 11 + j, 97, style, 10);
					setStringCellAndStyle(sheet, sheetstrength, 11 + j, 99, style, 10);
					setStringCellAndStyle(sheet, basethickness, 11 + j, 102, style, 11);
					// ??????1.2g????????????????????????????
					if (sheetstrength12.equals("1.2g")) {
						setStringCellAndStyle(sheet, "", 11 + j, 105, style, 10);
						setStringCellAndStyle(sheet, "", 11 + j, 108, style, 10);
						setStringCellAndStyle(sheet, "", 11 + j, 111, style, 10);
					} else {
						setStringCellAndStyle(sheet, CurrentSerie, 11 + j, 105, style, 10);
						setStringCellAndStyle(sheet, RecomWeldForce, 11 + j, 108, style, 10);
						setStringCellAndStyle(sheet, CurrentSerie, 11 + j, 111, style, 10);
					}
				}
			}

			shnum++;
		}

	}

	/*
	 * ????????????
	 */
	private Map<String, String[]> getHanJieParater(ArrayList hdlist) throws TCException {
		// TODO Auto-generated method stub
		Map<String, String[]> map = new HashMap<String, String[]>();
		if (hdlist != null && hdlist.size() > 0) {
			TCComponent[] tccs = new TCComponent[hdlist.size()];
			for (int i = 0; i < hdlist.size(); i++) {
				TCComponentBOMLine bl = (TCComponentBOMLine) hdlist.get(i);
				TCComponentItemRevision rev = bl.getItemRevision();
				tccs[i] = rev;
				String weldno = Util.getProperty(rev, "object_name");
			}
			String[] properties = { "object_name", "b8_CurrentSerie_Nissan", "b8_CurrentSerie_DFL", "b8_Current2",
					"b8_RecomWeldForce" };
			String[][] values = Util.getAllProperties(session, tccs, properties);
			for (int j = 0; j < values.length; j++) {
				String[] proper = values[j];
				map.put(proper[0], proper);
			}
		}
		return map;
	}

	/*
	 * ??????????????
	 */
	private void PoorPatternProcessing(XSSFWorkbook book, List assylist, boolean rLflag) {
		// TODO Auto-generated method stub
		ArrayList poorlist = new ArrayList();
		// ??????????sheet
		int sheetnum = 0;
		sheetnum = book.getNumberOfSheets();
		int sheetAtIndex = -1; // ??????????????
		for (int i = 0; i < sheetnum; i++) {
			String sheetname = book.getSheetName(i);
			if (sheetname.contains("??????")) {
				sheetAtIndex = i;
				break;
			}
		}
		if (sheetAtIndex == -1) {
			return;
		}
		int poornum = 0;// ????????????
		for (Map.Entry<String, String> entry : fymap.entrySet()) {
			String key = entry.getKey();
			int value = Integer.parseInt(entry.getValue());
			if (value > 1) {
				poornum++;
				List temp = new ArrayList();
				List afterName = new ArrayList(); // ??????????5??
				for (int i = 0; i < partlist.size(); i++) {
					String[] str = (String[]) partlist.get(i);
					if (key.equals(str[7])) {
						String[] station = new String[3];
						station[0] = str[1];
						if (str[1].length() > 5) {
							String afterno = str[1].substring(5);
							if (!afterName.contains(afterno)) {
								afterName.add(afterno);
							}
						} else {
							if (!afterName.contains(str[1])) {
								afterName.add(str[1]);
							}
						}
						station[1] = str[2];
						station[2] = Integer.toString(poornum);
						temp.add(station);
					}
				}
				// ????????????????????????????????????
				if (rLflag) {
					for (int j = 0; j < afterName.size(); j++) {
						String ApartNO = (String) afterName.get(j);
						List ttlist = new ArrayList();
						for (int k = 0; k < temp.size(); k++) {
							String[] val = (String[]) temp.get(k);
							if (val[0].length() > 5) {
								if (ApartNO.equals(val[0].substring(5))) {
									ttlist.add(val);
								}
							} else {
								if (ApartNO.equals(val[0])) {
									ttlist.add(val);
								}
							}
						}
						if (ttlist.size() == 2) {
							String[] str1 = (String[]) ttlist.get(0);
							String[] str2 = (String[]) ttlist.get(1);
							String[] str3 = new String[4];
							if (str1[0].length() > 4 && str2[0].length() > 4) {
								str3[0] = str1[0].substring(0, 5) + "/" + str2[0].substring(4, 5) + ApartNO;
							} else {
								str3[0] = str1[0];
							}
							if (str3[1] != null && str3[1].length() > 2) {
								str3[1] = str1[1].substring(0, str1[1].length() - 2) + "L/RH";
							} else {
								str3[1] = str1[1] + "L/RH";
							}

							str3[2] = str1[2];
							str3[3] = ApartNO;
							poorlist.add(str3);
						} else {
							if (ttlist.size() > 0) {
								String[] val = (String[]) temp.get(j);
								String[] strvalue = new String[4];
								if (val[0].length() > 5) {
									strvalue[0] = val[0].substring(0, 5);
									strvalue[1] = val[1];
									strvalue[2] = val[2];
									strvalue[3] = val[0].substring(5);
								} else {
									strvalue[0] = val[0];
									strvalue[1] = val[1];
									strvalue[2] = val[2];
									strvalue[3] = val[0];
								}

								poorlist.add(strvalue);
							}
						}
					}

				} else {
					for (int j = 0; j < temp.size(); j++) {
						String[] val = (String[]) temp.get(j);
						String[] strvalue = new String[4];
						if (val[0].length() > 5) {
							strvalue[0] = val[0].substring(0, 5);
							strvalue[1] = val[1];
							strvalue[2] = val[2];
							strvalue[3] = val[0].substring(5);
						} else {
							strvalue[0] = val[0];
							strvalue[1] = val[1];
							strvalue[2] = val[2];
							strvalue[3] = val[0];
						}

						poorlist.add(strvalue);
					}
				}
			}
		}
//		for (Map.Entry<String, String> entry : fymap.entrySet()) {
//			String key = entry.getKey();
//			int value = Integer.parseInt(entry.getValue());
//			if (value > 1) {
//				String partno = "";
//				String partname = "";
//				poornum++;
//
//				for (int i = 0; i < partlist.size(); i++) {
//					String[] str = (String[]) partlist.get(i);
//					if (key.equals(str[7])) {
//						String[] station = new String[3];
//						station[0] = str[1];
//						station[1] = str[2];
//						station[2] = Integer.toString(poornum);
//						poorlist.add(station);
//					}
//				}
//			}
//		}
		// ????????????????????????,??3??????????????
		int page = poornum / 3 + 1;

		// ??????????????????????sheet????????????????
		if (poornum % 3 == 0) {
			if (page > 1) {
				page = page - 1;
			}
		}

		// ????page????1????????????sheet??
		int index = sheetAtIndex + 1;

		// ????????????
		Font font = book.createFont();
		font.setColor((short) 12);// ????????
		// font.setFontName("????");
		font.setFontHeightInPoints((short) 14);
		XSSFCellStyle style = book.createCellStyle();
		style.setBorderBottom(XSSFCellStyle.BORDER_THIN); // ??????
		// style.setBorderLeft(XSSFCellStyle.BORDER_THIN);// ??????
		style.setBorderTop(XSSFCellStyle.BORDER_THIN);// ??????
		style.setBorderRight(XSSFCellStyle.BORDER_THIN);// ??????
		style.setVerticalAlignment(XSSFCellStyle.VERTICAL_CENTER);
		style.setAlignment(XSSFCellStyle.ALIGN_CENTER);
		style.setFont(font);

		XSSFCellStyle style2 = book.createCellStyle();
		style2.setBorderBottom(XSSFCellStyle.BORDER_THIN); // ??????
		style2.setBorderLeft(XSSFCellStyle.BORDER_MEDIUM);// ??????
		style2.setBorderTop(XSSFCellStyle.BORDER_THIN);// ??????
		style2.setBorderRight(XSSFCellStyle.BORDER_THIN);// ??????
		style2.setVerticalAlignment(XSSFCellStyle.VERTICAL_CENTER);
		style2.setAlignment(XSSFCellStyle.ALIGN_CENTER);
		style2.setFont(font);

		XSSFCellStyle style3 = book.createCellStyle();
		style3.setBorderBottom(XSSFCellStyle.BORDER_THIN); // ??????
		// style3.setBorderLeft(XSSFCellStyle.BORDER_THIN);// ??????
		style3.setBorderTop(XSSFCellStyle.BORDER_MEDIUM);// ??????
		style3.setBorderRight(XSSFCellStyle.BORDER_THIN);// ??????
		style3.setVerticalAlignment(XSSFCellStyle.VERTICAL_CENTER);
		style3.setAlignment(XSSFCellStyle.ALIGN_CENTER);
		style3.setFont(font);

		Font font2 = book.createFont();
		font2.setColor((short) 2);// ????????
		// font.setFontName("????");
		font2.setFontHeightInPoints((short) 14);
		XSSFCellStyle style00 = book.createCellStyle();
		style00.setBorderBottom(XSSFCellStyle.BORDER_THIN); // ??????
		// style.setBorderLeft(XSSFCellStyle.BORDER_THIN);// ??????
		style00.setBorderTop(XSSFCellStyle.BORDER_THIN);// ??????
		style00.setBorderRight(XSSFCellStyle.BORDER_THIN);// ??????
		style00.setVerticalAlignment(XSSFCellStyle.VERTICAL_CENTER);
		style00.setAlignment(XSSFCellStyle.ALIGN_CENTER);
		style00.setFont(font2);

		XSSFCellStyle style22 = book.createCellStyle();
		style22.setBorderBottom(XSSFCellStyle.BORDER_THIN); // ??????
		style22.setBorderLeft(XSSFCellStyle.BORDER_MEDIUM);// ??????
		style22.setBorderTop(XSSFCellStyle.BORDER_THIN);// ??????
		style22.setBorderRight(XSSFCellStyle.BORDER_THIN);// ??????
		style22.setVerticalAlignment(XSSFCellStyle.VERTICAL_CENTER);
		style22.setAlignment(XSSFCellStyle.ALIGN_CENTER);
		style22.setFont(font2);

		XSSFCellStyle style33 = book.createCellStyle();
		style33.setBorderBottom(XSSFCellStyle.BORDER_THIN); // ??????
		// style3.setBorderLeft(XSSFCellStyle.BORDER_THIN);// ??????
		style33.setBorderTop(XSSFCellStyle.BORDER_MEDIUM);// ??????
		style33.setBorderRight(XSSFCellStyle.BORDER_THIN);// ??????
		style33.setVerticalAlignment(XSSFCellStyle.VERTICAL_CENTER);
		style33.setAlignment(XSSFCellStyle.ALIGN_CENTER);
		style33.setFont(font2);

		/**************************************************/
		// ????????????????????????????????????????????????
		List assynameList = new ArrayList();// ????????????
		List assyList = new ArrayList();// ????????????
		if (updateflag) {
			int gcnum = 0;
			for (int i = 0; i < sheetnum; i++) {
				String sheetname = book.getSheetName(i);
				if (sheetname.contains("??????")) {
					gcnum++;
				}
			}
			// ????sheet????????????????????????????
			index = sheetAtIndex + page;

			XSSFCell cell;
			XSSFRow row;
			// ??????????sheet????????????????????????????????????
			for (int i = sheetAtIndex; i < sheetAtIndex + gcnum; i++) {
				XSSFSheet sheet = book.getSheetAt(i);
				row = sheet.getRow(6);
				if (row != null) {
					cell = row.getCell(9);
					if (cell != null) {
						String assyname = convertCellValueToString(cell, Cell.CELL_TYPE_STRING);
						if (assyname != null && !assyname.isEmpty()) {
							assynameList.add(assyname);
						}
					}
					cell = row.getCell(46);
					if (cell != null) {
						String assyname = convertCellValueToString(cell, Cell.CELL_TYPE_STRING);
						if (assyname != null && !assyname.isEmpty()) {
							assynameList.add(assyname);
						}
					}
					cell = row.getCell(84);
					if (cell != null) {
						String assyname = convertCellValueToString(cell, Cell.CELL_TYPE_STRING);
						if (assyname != null && !assyname.isEmpty()) {
							assynameList.add(assyname);
						}
					}
				}
				// ????????
				setStringCellAndStyle(sheet, "", 6, 9, style3, Cell.CELL_TYPE_STRING);// ????????????
				setStringCellAndStyle(sheet, "", 6, 46, style3, Cell.CELL_TYPE_STRING);// ????????????
				setStringCellAndStyle(sheet, "", 6, 84, style3, Cell.CELL_TYPE_STRING);// ????????????

				for (int j = 0; j < 6; j++) {
					row = sheet.getRow(34 + j * 2);
					if (row != null) {
						cell = row.getCell(1);
						if (cell != null) {
							String preassy = convertCellValueToString(cell, Cell.CELL_TYPE_STRING);
							cell = row.getCell(7);
							String sufxxassy = convertCellValueToString(cell, Cell.CELL_TYPE_STRING);
							if ((preassy != null && !preassy.isEmpty())
									|| (sufxxassy != null && !sufxxassy.isEmpty())) {
								String asstno = preassy.trim() + sufxxassy.trim();
								assyList.add(asstno);
							}
						}
						cell = row.getCell(38);
						if (cell != null) {
							String preassy = convertCellValueToString(cell, Cell.CELL_TYPE_STRING);
							cell = row.getCell(44);
							String sufxxassy = convertCellValueToString(cell, Cell.CELL_TYPE_STRING);
							if ((preassy != null && !preassy.isEmpty())
									|| (sufxxassy != null && !sufxxassy.isEmpty())) {
								String asstno = preassy.trim() + sufxxassy.trim();
								assyList.add(asstno);
							}
						}
						cell = row.getCell(76);
						if (cell != null) {
							String preassy = convertCellValueToString(cell, Cell.CELL_TYPE_STRING);
							cell = row.getCell(82);
							String sufxxassy = convertCellValueToString(cell, Cell.CELL_TYPE_STRING);
							if ((preassy != null && !preassy.isEmpty())
									|| (sufxxassy != null && !sufxxassy.isEmpty())) {
								String asstno = preassy.trim() + sufxxassy.trim();
								assyList.add(asstno);
							}
						}
					}
					setStringCellAndStyle(sheet, "", 34 + j * 2, 1, style2, Cell.CELL_TYPE_STRING);// ????????????????
					setStringCellAndStyle(sheet, "", 34 + j * 2, 7, style, Cell.CELL_TYPE_STRING);// ????????????????
					setStringCellAndStyle(sheet, "", 34 + j * 2, 38, style2, Cell.CELL_TYPE_STRING);// ????????????????
					setStringCellAndStyle(sheet, "", 34 + j * 2, 44, style, Cell.CELL_TYPE_STRING);// ????????????????
					setStringCellAndStyle(sheet, "", 34 + j * 2, 76, style2, Cell.CELL_TYPE_STRING);// ????????????????
					setStringCellAndStyle(sheet, "", 34 + j * 2, 82, style, Cell.CELL_TYPE_STRING);// ????????????????
				}
			}
			if (page > gcnum) {
				for (int i = 1; i < page - gcnum; i++) {
					XSSFSheet newsheet = book.cloneSheet(sheetAtIndex);
					book.setSheetOrder(newsheet.getSheetName(), index);
					index++;
				}
			}
		} else {
			if (page > 1) {
				for (int i = 1; i < page; i++) {
					XSSFSheet newsheet = book.cloneSheet(sheetAtIndex);
					book.setSheetOrder(newsheet.getSheetName(), index);
					index++;
				}
			}
		}
		/**************************************************/

		int shnum = 0;
		for (int i = sheetAtIndex; i < index; i++) {
			XSSFSheet sheet = book.getSheetAt(i);
			if (i == index - 1) {
				for (int j = 0; j + 3 * shnum < poornum; j++) {
					String partname = "";
					int rownum = 0;
					for (int k = 0; k < poorlist.size(); k++) {
						String[] str = (String[]) poorlist.get(k);
						if (j + 1 + 3 * shnum == Integer.parseInt(str[2])) {
							partname = str[1];
							String prename = str[0];
							String aftername = str[3];
//							if (str[0].length() > 5) {
//								prename = str[0].substring(0, 5);
//								aftername = str[0].substring(5).trim();
//							} else {
//								prename = str[0];
//								aftername = str[0];
//							}
							if ((j + 1 + 3 * shnum) % 3 == 1) {
								if (updateflag && assyList != null) {
									String allassy = prename.trim() + aftername.trim();
									if (!assyList.contains(allassy)) {
										setStringCellAndStyle2(sheet, prename, 34 + rownum * 2, 1, style22,
												Cell.CELL_TYPE_STRING);// ????????????????
										setStringCellAndStyle2(sheet, aftername, 34 + rownum * 2, 7, style00,
												Cell.CELL_TYPE_STRING);// ????????????????
									} else {
										setStringCellAndStyle2(sheet, prename, 34 + rownum * 2, 1, style2,
												Cell.CELL_TYPE_STRING);// ????????????????
										setStringCellAndStyle2(sheet, aftername, 34 + rownum * 2, 7, style,
												Cell.CELL_TYPE_STRING);// ????????????????
									}
								} else {
									setStringCellAndStyle2(sheet, prename, 34 + rownum * 2, 1, style2,
											Cell.CELL_TYPE_STRING);// ????????????????
									setStringCellAndStyle2(sheet, aftername, 34 + rownum * 2, 7, style,
											Cell.CELL_TYPE_STRING);// ????????????????
								}

							} else if ((j + 1 + 3 * shnum) % 3 == 2) {
								if (updateflag && assyList != null) {
									String allassy = prename.trim() + aftername.trim();
									if (!assyList.contains(allassy)) {
										setStringCellAndStyle2(sheet, prename, 34 + rownum * 2, 38, style22,
												Cell.CELL_TYPE_STRING);// ????????????????
										setStringCellAndStyle2(sheet, aftername, 34 + rownum * 2, 44, style00,
												Cell.CELL_TYPE_STRING);// ????????????????
									} else {
										setStringCellAndStyle2(sheet, prename, 34 + rownum * 2, 38, style2,
												Cell.CELL_TYPE_STRING);// ????????????????
										setStringCellAndStyle2(sheet, aftername, 34 + rownum * 2, 44, style,
												Cell.CELL_TYPE_STRING);// ????????????????
									}
								} else {
									setStringCellAndStyle2(sheet, prename, 34 + rownum * 2, 38, style2,
											Cell.CELL_TYPE_STRING);// ????????????????
									setStringCellAndStyle2(sheet, aftername, 34 + rownum * 2, 44, style,
											Cell.CELL_TYPE_STRING);// ????????????????
								}

							} else {
								if (updateflag && assyList != null) {
									String allassy = prename.trim() + aftername.trim();
									if (!assyList.contains(allassy)) {
										setStringCellAndStyle2(sheet, prename, 34 + rownum * 2, 76, style22,
												Cell.CELL_TYPE_STRING);// ????????????????
										setStringCellAndStyle2(sheet, aftername, 34 + rownum * 2, 82, style00,
												Cell.CELL_TYPE_STRING);// ????????????????
									} else {
										setStringCellAndStyle2(sheet, prename, 34 + rownum * 2, 76, style2,
												Cell.CELL_TYPE_STRING);// ????????????????
										setStringCellAndStyle2(sheet, aftername, 34 + rownum * 2, 82, style,
												Cell.CELL_TYPE_STRING);// ????????????????
									}
								} else {
									setStringCellAndStyle2(sheet, prename, 34 + rownum * 2, 76, style2,
											Cell.CELL_TYPE_STRING);// ????????????????
									setStringCellAndStyle2(sheet, aftername, 34 + rownum * 2, 82, style,
											Cell.CELL_TYPE_STRING);// ????????????????
								}

							}
							rownum++;
						}
					}
					if ((j + 1 + 3 * shnum) % 3 == 1) {
						if (updateflag && assynameList != null) {
							if (!assynameList.contains(partname)) {
								setStringCellAndStyle2(sheet, partname, 6, 9, style33, Cell.CELL_TYPE_STRING);// ????????????
							} else {
								setStringCellAndStyle2(sheet, partname, 6, 9, style3, Cell.CELL_TYPE_STRING);// ????????????
							}
						} else {
							setStringCellAndStyle2(sheet, partname, 6, 9, style3, Cell.CELL_TYPE_STRING);// ????????????
						}

					} else if ((j + 1 + 3 * shnum) % 3 == 2) {
						if (updateflag && assynameList != null) {
							if (!assynameList.contains(partname)) {
								setStringCellAndStyle2(sheet, partname, 6, 46, style33, Cell.CELL_TYPE_STRING);// ????????????
							} else {
								setStringCellAndStyle2(sheet, partname, 6, 46, style3, Cell.CELL_TYPE_STRING);// ????????????
							}
						} else {
							setStringCellAndStyle2(sheet, partname, 6, 46, style3, Cell.CELL_TYPE_STRING);// ????????????
						}
					} else {
						if (updateflag && assynameList != null) {
							if (!assynameList.contains(partname)) {
								setStringCellAndStyle2(sheet, partname, 6, 84, style33, Cell.CELL_TYPE_STRING);// ????????????
							} else {
								setStringCellAndStyle2(sheet, partname, 6, 84, style3, Cell.CELL_TYPE_STRING);// ????????????
							}
						} else {
							setStringCellAndStyle2(sheet, partname, 6, 84, style3, Cell.CELL_TYPE_STRING);// ????????????
						}
					}
				}
			} else {
				for (int j = 0; j + 3 * shnum < 3 + 3 * shnum; j++) {
					String partname = "";
					int rownum = 0;
					for (int k = 0; k < poorlist.size(); k++) {
						String[] str = (String[]) poorlist.get(k);
						if (j + 1 + 3 * shnum == Integer.parseInt(str[2])) {
							partname = str[1];
							String prename = str[0];
							String aftername = str[3];
//							if (str[0].length() > 5) {
//								prename = str[0].substring(0, 5);
//								aftername = str[0].substring(5).trim();
//							} else {
//								prename = str[0];
//								aftername = str[0];
//							}
							if ((j + 1 + 3 * shnum) % 3 == 1) {
								if (updateflag && assyList != null) {
									String allassy = prename.trim() + aftername.trim();
									if (!assyList.contains(allassy)) {
										setStringCellAndStyle2(sheet, prename, 34 + rownum * 2, 1, style22,
												Cell.CELL_TYPE_STRING);// ????????????????
										setStringCellAndStyle2(sheet, aftername, 34 + rownum * 2, 7, style00,
												Cell.CELL_TYPE_STRING);// ????????????????
									} else {
										setStringCellAndStyle2(sheet, prename, 34 + rownum * 2, 1, style2,
												Cell.CELL_TYPE_STRING);// ????????????????
										setStringCellAndStyle2(sheet, aftername, 34 + rownum * 2, 7, style,
												Cell.CELL_TYPE_STRING);// ????????????????
									}
								} else {
									setStringCellAndStyle2(sheet, prename, 34 + rownum * 2, 1, style2,
											Cell.CELL_TYPE_STRING);// ????????????????
									setStringCellAndStyle2(sheet, aftername, 34 + rownum * 2, 7, style,
											Cell.CELL_TYPE_STRING);// ????????????????
								}

							} else if ((j + 1 + 3 * shnum) % 3 == 2) {
								if (updateflag && assyList != null) {
									String allassy = prename.trim() + aftername.trim();
									if (!assyList.contains(allassy)) {
										setStringCellAndStyle2(sheet, prename, 34 + rownum * 2, 38, style22,
												Cell.CELL_TYPE_STRING);// ????????????????
										setStringCellAndStyle2(sheet, aftername, 34 + rownum * 2, 44, style00,
												Cell.CELL_TYPE_STRING);// ????????????????
									} else {
										setStringCellAndStyle2(sheet, prename, 34 + rownum * 2, 38, style2,
												Cell.CELL_TYPE_STRING);// ????????????????
										setStringCellAndStyle2(sheet, aftername, 34 + rownum * 2, 44, style,
												Cell.CELL_TYPE_STRING);// ????????????????
									}
								} else {
									setStringCellAndStyle2(sheet, prename, 34 + rownum * 2, 38, style2,
											Cell.CELL_TYPE_STRING);// ????????????????
									setStringCellAndStyle2(sheet, aftername, 34 + rownum * 2, 44, style,
											Cell.CELL_TYPE_STRING);// ????????????????
								}

							} else {
								if (updateflag && assyList != null) {
									String allassy = prename.trim() + aftername.trim();
									if (!assyList.contains(allassy)) {
										setStringCellAndStyle2(sheet, prename, 34 + rownum * 2, 76, style22,
												Cell.CELL_TYPE_STRING);// ????????????????
										setStringCellAndStyle2(sheet, aftername, 34 + rownum * 2, 82, style00,
												Cell.CELL_TYPE_STRING);// ????????????????
									} else {
										setStringCellAndStyle2(sheet, prename, 34 + rownum * 2, 76, style2,
												Cell.CELL_TYPE_STRING);// ????????????????
										setStringCellAndStyle2(sheet, aftername, 34 + rownum * 2, 82, style,
												Cell.CELL_TYPE_STRING);// ????????????????
									}
								} else {
									setStringCellAndStyle2(sheet, prename, 34 + rownum * 2, 76, style2,
											Cell.CELL_TYPE_STRING);// ????????????????
									setStringCellAndStyle2(sheet, aftername, 34 + rownum * 2, 82, style,
											Cell.CELL_TYPE_STRING);// ????????????????
								}

							}
							rownum++;
						}
					}
					if ((j + 1 + 3 * shnum) % 3 == 1) {
						if (updateflag && assynameList != null) {
							if (!assynameList.contains(partname)) {
								setStringCellAndStyle2(sheet, partname, 6, 9, style33, Cell.CELL_TYPE_STRING);// ????????????
							} else {
								setStringCellAndStyle2(sheet, partname, 6, 9, style3, Cell.CELL_TYPE_STRING);// ????????????
							}
						} else {
							setStringCellAndStyle2(sheet, partname, 6, 9, style3, Cell.CELL_TYPE_STRING);// ????????????
						}

					} else if ((j + 1 + 3 * shnum) % 3 == 2) {
						if (updateflag && assynameList != null) {
							if (!assynameList.contains(partname)) {
								setStringCellAndStyle2(sheet, partname, 6, 46, style33, Cell.CELL_TYPE_STRING);// ????????????
							} else {
								setStringCellAndStyle2(sheet, partname, 6, 46, style3, Cell.CELL_TYPE_STRING);// ????????????
							}
						} else {
							setStringCellAndStyle2(sheet, partname, 6, 46, style3, Cell.CELL_TYPE_STRING);// ????????????
						}
					} else {
						if (updateflag && assynameList != null) {
							if (!assynameList.contains(partname)) {
								setStringCellAndStyle2(sheet, partname, 6, 84, style33, Cell.CELL_TYPE_STRING);// ????????????
							} else {
								setStringCellAndStyle2(sheet, partname, 6, 84, style3, Cell.CELL_TYPE_STRING);// ????????????
							}
						} else {
							setStringCellAndStyle2(sheet, partname, 6, 84, style3, Cell.CELL_TYPE_STRING);// ????????????
						}
					}
				}
			}
			shnum++;
		}

	}

	/*
	 * ??????????????
	 */
	private void CompositionChartProcessing(XSSFWorkbook book, List assylist, String assyname, boolean rLflag,
			Map<String, File> piclist) {
		// TODO Auto-generated method stub
//		if (!updateflag) 
//		{
//			for(int i=0;i<deletelist.size();i++)
//			{
//				if(deletelist.get(i).toString().contains("??????"))
//				{
//					return;
//				}
//			}			
//		}	
		ArrayList complist = new ArrayList();
		// ??????????sheet
		int sheetnum = 0;
		sheetnum = book.getNumberOfSheets();
		int sheetAtIndex = -1; // ??????????????
		for (int i = 0; i < sheetnum; i++) {
			String sheetname = book.getSheetName(i);
			if (sheetname.contains("??????")) {
				sheetAtIndex = i;
				break;
			}
		}
		if (sheetAtIndex == -1) {
			return;
		}
		//????????????
//		Map<String,String> partTonum = new HashMap<String,String>();
//		for (int i = 0; i < partlist.size(); i++) 
//		{
//			String[] str = (String[]) partlist.get(i);
//			if(partlist.contains(str[1]))
//			{
//				partTonum.put(str[1], str[2]);
//			}
//		}
		// ????????????????????????????????????????????????????
		for (Map.Entry<String, String> entry : fymap.entrySet()) {
			String key = entry.getKey();
			// ??????????????????????
			List tempList = new ArrayList();
			List afterName = new ArrayList(); // ??????????5??
			for (int i = 0; i < partlist.size(); i++) {
				String[] str = (String[]) partlist.get(i);
				if (key.equals(str[7])) {
					tempList.add(str);
					if (str[1] != null && str[1].length() > 5) {
						String afterno = str[1].substring(5);
						if (!afterName.contains(afterno)) {
							afterName.add(afterno);
						}
					} else {
						if (!afterName.contains(str[1])) {
							afterName.add(str[1]);
						}
					}
				}
			}
			// ??????????????????????????5????????????????????
			if (rLflag) {
				for (int j = 0; j < afterName.size(); j++) {
					String ApartNO = (String) afterName.get(j);
					if(ApartNO == null)
					{
						ApartNO = "";
					}
					List ttlist = new ArrayList();
					for (int k = 0; k < tempList.size(); k++) {
						String[] val = (String[]) tempList.get(k);
						if (val[1] != null && val[1].length() > 5) {
							if (ApartNO.equals(val[1].substring(5))) {
								ttlist.add(val);
							}
						} else {
							if (ApartNO.equals(val[1])) {
								ttlist.add(val);
							}
						}
					}
					if (ttlist.size() == 2) {
						String[] str1 = (String[]) ttlist.get(0);
						String[] str2 = (String[]) ttlist.get(1);
						String[] str3 = new String[3];
						if (str1[1] != null && str1[1].length() > 5 && str2[1] != null && str2[1].length() > 5) {
							str3[0] = str1[1].substring(0, 5) + "/" + str2[1].substring(4, 5) + ApartNO;
						} else {
							str3[0] = str1[1] + "/" + str2[1];
						}
						if (str1[2] != null && str1[2].length() > 2) {
							str3[1] = str1[2].substring(0, str1[2].length() - 2) + "L/RH";
						} else {
							str3[1] = str1[2] + "L/RH";
						}

						str3[2] = str1[2];
						//??????????????????????????????????????????????
						if(str1[2].equals(str2[2]))
						{
							String[] station = new String[3];
							station[0] = str1[7] + " " + str1[1] + " " + str1[2];
							station[1] = str1[6];
							station[2] = str1[8];
							complist.add(station);
							
							String[] station2 = new String[3];
							station2[0] = str2[7] + " " + str2[1] + " " + str2[2];
							station2[1] = str2[6];
							station2[2] = str2[8];
							complist.add(station2);
						}
						else
						{
							String[] station = new String[3];
							station[0] = str1[7] + " " + str3[0] + " " + str3[1];
							station[1] = str1[6];
							station[2] = str1[8];
							complist.add(station);
						}
						
					} else {
						if (ttlist.size() > 0) {
							for (int m = 0; m < ttlist.size(); m++) {
								String[] val = (String[]) ttlist.get(m);
								String[] station = new String[3];
								station[0] = val[7] + " " + val[1] + " " + val[2];
								station[1] = val[6];
								station[2] = val[8];
								complist.add(station);
							}
						}
					}
				}

			} else {
				for (int j = 0; j < tempList.size(); j++) {
					String[] val = (String[]) tempList.get(j);
					String[] station = new String[3];
					station[0] = val[7] + " " + val[1] + " " + val[2];
					station[1] = val[6];
					station[2] = val[8];
					complist.add(station);
				}
			}
		}

//		for (Map.Entry<String, String> entry : fymap.entrySet()) {
//			String key = entry.getKey();
//			String stationname = "";
//			String template = "";
//			String type = "";
//			String[] station = new String[2];
//			for (int i = 0; i < partlist.size(); i++) {
//				String[] str = (String[]) partlist.get(i);
//				if (key.equals(str[7])) {
//					if (template.isEmpty()) {
//						template = str[1];
//					} else {
//						if (str[1].length() > 5) {
//							template = template + "/" + str[1].substring(5).trim();
//						} else {
//							template = template + "/" + str[1];
//						}
//					}
//					stationname = template + " " + str[2];
//					type = str[6];
//				}
//			}
//			station[0] = stationname;
//			station[1] = type;
//			template = "";
//			complist.add(station);
//		}

		// ????????????????????????????,????15??????????????????
		// int sum = fymap.size();
		int sum = complist.size();
		int page = sum / 15 + 1;
		// ????page????1??????????????

		// ????????????
		// ????????????
		Font font = book.createFont();
		font.setColor((short) 12);
		font.setFontName("????");
		XSSFCellStyle style = book.createCellStyle();
		style.setFont(font);
		style.setBorderBottom(CellStyle.BORDER_DOTTED); // ???????? BORDER_HAIR  
		style.setBorderLeft(CellStyle.BORDER_THIN); // ???????? BORDER_DOTTED
		style.setBorderRight(CellStyle.BORDER_DOTTED); // ????????
		style.setBorderTop(CellStyle.BORDER_DOTTED); // ????????
		

		XSSFCellStyle style2 = book.createCellStyle();
		style2.setFont(font);
		style2.setBorderBottom(CellStyle.BORDER_DOUBLE); // ???????? 
		style2.setBorderLeft(CellStyle.BORDER_THIN); // ????????
		style2.setBorderRight(CellStyle.BORDER_DOUBLE); // ????????
		style2.setBorderTop(CellStyle.BORDER_DOUBLE); // ????????

		XSSFCellStyle style3 = book.createCellStyle();
		style3.setFont(font);
		style3.setBorderBottom(CellStyle.BORDER_MEDIUM); // ????????
		style3.setBorderLeft(CellStyle.BORDER_THIN); // ????????
		style3.setBorderRight(CellStyle.BORDER_MEDIUM); // ????????
		style3.setBorderTop(CellStyle.BORDER_MEDIUM); // ????????

		XSSFCellStyle style4 = book.createCellStyle();
		style4.setFont(font);
		style4.setBorderBottom(CellStyle.BORDER_NONE); // ??????
		style4.setBorderLeft(CellStyle.BORDER_NONE); // ??????
		style4.setBorderRight(CellStyle.BORDER_NONE); // ??????
		style4.setBorderTop(CellStyle.BORDER_NONE); // ??????

		XSSFCellStyle style5 = book.createCellStyle();
		style5.setFont(font);
		style5.setBorderLeft(CellStyle.BORDER_THIN); // ??????

		Font font2 = book.createFont();
		font2.setColor((short) 2);
		font2.setFontName("????");
		XSSFCellStyle style00 = book.createCellStyle();
		style00.setFont(font2);
		style00.setBorderBottom(CellStyle.BORDER_DOTTED); // ????????
		style00.setBorderLeft(CellStyle.BORDER_THIN); // ????????
		style00.setBorderRight(CellStyle.BORDER_DOTTED); // ????????
		style00.setBorderTop(CellStyle.BORDER_DOTTED); // ????????

		XSSFCellStyle style22 = book.createCellStyle();
		style22.setFont(font2);
		style22.setBorderBottom(CellStyle.BORDER_DOUBLE); // ????????
		style22.setBorderLeft(CellStyle.BORDER_THIN); // ????????
		style22.setBorderRight(CellStyle.BORDER_DOUBLE); // ????????
		style22.setBorderTop(CellStyle.BORDER_DOUBLE); // ????????

		XSSFCellStyle style33 = book.createCellStyle();
		style33.setFont(font2);
		style33.setBorderBottom(CellStyle.BORDER_MEDIUM); // ????????
		style33.setBorderLeft(CellStyle.BORDER_THIN); // ????????
		style33.setBorderRight(CellStyle.BORDER_MEDIUM); // ????????
		style33.setBorderTop(CellStyle.BORDER_MEDIUM); // ????????

		XSSFCellStyle style44 = book.createCellStyle();
		style44.setFont(font2);
		style44.setBorderBottom(CellStyle.BORDER_NONE); // ??????
		style44.setBorderLeft(CellStyle.BORDER_NONE); // ??????
		style44.setBorderRight(CellStyle.BORDER_NONE); // ??????
		style44.setBorderTop(CellStyle.BORDER_NONE); // ??????

		int shnum = 0;

		// ????????????
		String procStation = "";
		String tempstation = "";
//		if (assynos != null) {
//			for (int i = 0; i < assynos.length; i++) {
//				if (i == 0) {
//					tempstation = assynos[i];
//				} else {
//					if (assynos[i].length() > 5) {
//						tempstation = tempstation + "/" + assynos[i].substring(5).trim();
//					} else {
//						tempstation = tempstation + "/" + assynos[i];
//					}
//				}
//			}
//		}
		if (assylist != null) {
			List aftertemp = new ArrayList();
			String fronttemp = "";
			String backtemp = "";
			for (int i = 0; i < assylist.size(); i++) {
				String strVal = (String) assylist.get(i);
				if (i == 0) {
					if (strVal != null && strVal.length() >= 5) {
						fronttemp = strVal.substring(0, 5);
						backtemp = strVal.substring(5).trim();
					} else {
						fronttemp = strVal;
						backtemp = "";
					}

				} else {
					if (strVal != null && strVal.length() >= 5) {
						if (!fronttemp.equals(strVal.substring(0, 5))) {
							fronttemp = fronttemp + "/" + strVal.substring(4, 5);
						}
						if (!backtemp.equals(strVal.substring(5))) {
							backtemp = backtemp + "/" + strVal.substring(5).trim();
						}
					} else {
						fronttemp = fronttemp + strVal;
						backtemp = "";
					}
				}
			}
			tempstation = fronttemp + " " + backtemp;
			System.out.println("tempstation:" + tempstation);
		}

		if (rLflag) {
			procStation = tempstation + " " + assyname.replace("LH", "").replace("RH", "") + " " + "L/RH";
		} else {
			procStation = tempstation + " " + assyname;
		}
		System.out.println("procStation:" + procStation);
		// ????????????????????????
		if (assylist == null || assylist.size() < 1) {
			procStation = "";
		}

		XSSFSheet sheet = book.getSheetAt(sheetAtIndex);
		/***********************************************/
		// ????????????????????????????????????????????????
		String oldproc = "FLAG";
		List oldcomplist = new ArrayList();// partlist ????
		if (updateflag) {
			XSSFRow row;
			XSSFCell cell;
			row = sheet.getRow(8);
			if (row != null) {
				cell = row.getCell(3);
				if (cell != null) {
					oldproc = convertCellValueToString(cell, Cell.CELL_TYPE_STRING);
				}
			}
			setStringCellAndStyle(sheet, "", 8, 3, style4, Cell.CELL_TYPE_STRING);
			for (int j = 0; j < 30; j++) {
				row = sheet.getRow(10 + j);
				if (row != null) {
					cell = row.getCell(6);
					if (cell != null) {
						String partname = convertCellValueToString(cell, Cell.CELL_TYPE_STRING);
						if (partname != null && !partname.isEmpty()) {
							oldcomplist.add(partname);
						}
					}
				}

				setStringCellAndStyle(sheet, "", 10 + j, 6, style4, Cell.CELL_TYPE_STRING);
				for (int n = 7; n < 32; n++) {
					setStringCellAndStyle(sheet, "", 10 + j, n, style4, Cell.CELL_TYPE_STRING);
				}
			}
		}
		/***********************************************/
		CellRangeAddress region = null;
		if (!oldproc.equals("FLAG") && !oldproc.equals(procStation.trim())) {
			setStringCellAndStyle2(sheet, procStation.trim(), 8, 3, style22, Cell.CELL_TYPE_STRING);
		} else {
			setStringCellAndStyle2(sheet, procStation.trim(), 8, 3, style2, Cell.CELL_TYPE_STRING);
		}
		// ??????????????????????????????
		for (int n = 4; n < 32; n++) {
			setStringCellAndStyle2(sheet, "", 8, n, style2, Cell.CELL_TYPE_STRING);
		}
//		sheet.removeMergedRegion(getMergedRegionIndex(sheet, 8, 3));
//		region = new CellRangeAddress(8, 8, (short) 3, (short) 31);
//		sheet.addMergedRegion(region);
		if (page > 1) {
			setStringCellAndStyle2(sheet, "", 9, 6, style5, Cell.CELL_TYPE_STRING);
			for (int j = 0; j < complist.size(); j++) {
				String[] str = (String[]) complist.get(j);
				if (str[2] != null && str[2].equals("Stamping")) {
					if (oldcomplist != null && updateflag) {
						if (!oldcomplist.contains(str[0].trim())) {
							setStringCellAndStyle2(sheet, str[0].trim(), 10 + j, 6, style33, Cell.CELL_TYPE_STRING);
						} else {
							setStringCellAndStyle2(sheet, str[0].trim(), 10 + j, 6, style3, Cell.CELL_TYPE_STRING);
						}
					} else {
						setStringCellAndStyle2(sheet, str[0].trim(), 10 + j, 6, style3, Cell.CELL_TYPE_STRING);
					}
					for (int n = 7; n < 32; n++) {
						setStringCellAndStyle2(sheet, "", 10 + j, n, style3, Cell.CELL_TYPE_STRING);
					}
				} else if (str[2] != null && str[2].equals("Purchase")) {
					if (oldcomplist != null && updateflag) {
						if (!oldcomplist.contains(str[0].trim())) {
							setStringCellAndStyle2(sheet, str[0].trim(), 10 + j, 6, style00, Cell.CELL_TYPE_STRING);
						} else {
							setStringCellAndStyle2(sheet, str[0].trim(), 10 + j, 6, style, Cell.CELL_TYPE_STRING);
						}
					} else {
						setStringCellAndStyle2(sheet, str[0].trim(), 10 + j, 6, style, Cell.CELL_TYPE_STRING);
					}

					for (int n = 7; n < 32; n++) {
						setStringCellAndStyle2(sheet, "", 10 + j, n, style, Cell.CELL_TYPE_STRING);
					}
				} else {
					if (oldcomplist != null && updateflag) {
						if (!oldcomplist.contains(str[0].trim())) {
							setStringCellAndStyle2(sheet, str[0].trim(), 10 + j, 6, style22, Cell.CELL_TYPE_STRING);
						} else {
							setStringCellAndStyle2(sheet, str[0].trim(), 10 + j, 6, style2, Cell.CELL_TYPE_STRING);
						}
					} else {
						setStringCellAndStyle2(sheet, str[0].trim(), 10 + j, 6, style2, Cell.CELL_TYPE_STRING);
					}

					for (int n = 7; n < 32; n++) {
						setStringCellAndStyle2(sheet, "", 10 + j, n, style2, Cell.CELL_TYPE_STRING);
					}
				}
				// sheet.removeMergedRegion(getMergedRegionIndex(sheet, 10 + j, 6));
				region = new CellRangeAddress(10 + j, 10 + j, (short) 6, (short) 31);
				if (updateflag) {
					int mr1 = getMergedRegionIndex(sheet, 10 + j, 6);
					System.out.println("mr1: " + mr1);
					if (mr1 == -1) {
						sheet.addMergedRegion(region);
					}
				} else {
					sheet.addMergedRegion(region);
				}

			}
		} else {
			for (int j = 0; j < complist.size(); j++) {
				String[] str = (String[]) complist.get(j);
				setStringCellAndStyle2(sheet, "", 10 + j * 2 - 1, 6, style5, Cell.CELL_TYPE_STRING);
				if (str[2] != null && str[2].equals("Stamping")) {
					if (oldcomplist != null && updateflag) {
						if (!oldcomplist.contains(str[0].trim())) {
							setStringCellAndStyle2(sheet, str[0].trim(), 10 + j * 2, 6, style33, Cell.CELL_TYPE_STRING);
						} else {
							setStringCellAndStyle2(sheet, str[0].trim(), 10 + j * 2, 6, style3, Cell.CELL_TYPE_STRING);
						}
					} else {
						setStringCellAndStyle2(sheet, str[0].trim(), 10 + j * 2, 6, style3, Cell.CELL_TYPE_STRING);
					}

					for (int n = 7; n < 32; n++) {
						setStringCellAndStyle2(sheet, "", 10 + j * 2, n, style3, Cell.CELL_TYPE_STRING);
					}
				} else if (str[2] != null && str[2].equals("Purchase")) {
					if (oldcomplist != null && updateflag) {
						if (!oldcomplist.contains(str[0].trim())) {
							setStringCellAndStyle2(sheet, str[0].trim(), 10 + j * 2, 6, style00, Cell.CELL_TYPE_STRING);
						} else {
							setStringCellAndStyle2(sheet, str[0].trim(), 10 + j * 2, 6, style, Cell.CELL_TYPE_STRING);
						}
					} else {
						setStringCellAndStyle2(sheet, str[0].trim(), 10 + j * 2, 6, style, Cell.CELL_TYPE_STRING);
					}

					for (int n = 7; n < 32; n++) {
						setStringCellAndStyle2(sheet, "", 10 + j * 2, n, style, Cell.CELL_TYPE_STRING);
					}
				} else {
					if (oldcomplist != null && updateflag) {
						if (!oldcomplist.contains(str[0].trim())) {
							setStringCellAndStyle2(sheet, str[0].trim(), 10 + j * 2, 6, style22, Cell.CELL_TYPE_STRING);
						} else {
							setStringCellAndStyle2(sheet, str[0].trim(), 10 + j * 2, 6, style2, Cell.CELL_TYPE_STRING);
						}
					} else {
						setStringCellAndStyle2(sheet, str[0].trim(), 10 + j * 2, 6, style2, Cell.CELL_TYPE_STRING);
					}
					for (int n = 7; n < 32; n++) {
						setStringCellAndStyle2(sheet, "", 10 + j * 2, n, style2, Cell.CELL_TYPE_STRING);
					}
				}
				// sheet.removeMergedRegion(getMergedRegionIndex(sheet, 10 + j * 2, 6));
				region = new CellRangeAddress(10 + j * 2, 10 + j * 2, (short) 6, (short) 31);
				if (updateflag) {
					int mr1 = getMergedRegionIndex(sheet, 10 + j * 2, 6);
					System.out.println("mr1: " + mr1);
					if (mr1 == -1) {
						sheet.addMergedRegion(region);
					}
				} else {
					sheet.addMergedRegion(region);
				}
			}
		}

		// ??????????sheet??
		if (piclist != null && piclist.size() > 0) {

			// ????????????????????HSSFPatriarch, ????sheet????????????
			XSSFDrawing hssfPatriarch = (XSSFDrawing) sheet.createDrawingPatriarch();
			XSSFClientAnchor anchor1 = null;
			XSSFRichTextString strValue = new XSSFRichTextString();
			int count = 0;
			int rowindex = 0;
			int colindex = 0;

			for (Entry<String, File> entry : piclist.entrySet()) {
				if ((count + 1) % 3 == 1) {
					rowindex = 0;
				} else if ((count + 1) % 3 == 2) {
					rowindex = 1;
				} else {
					rowindex = 2;
				}
				colindex = (count) / 3;
				count++;
				// String objectname = entry.getKey().replace("/", " ");
				File file = entry.getValue();
				BufferedImage image = null;
				try {
					image = ImageIO.read(file);
				} catch (IOException e) {
					// TODO Auto-generated catch block
					e.printStackTrace();
				}
				int width = image.getWidth();
				int hight = image.getHeight();
				double diff = width * 1.0 / hight;
				int h = 13;
				int w = (int) (h * diff);

				writepicturetosheet(book, sheet, image, 5 + rowindex * 14, 37 + colindex * 25, 18 + rowindex * 14,
						(37 + w) + colindex * 25);
			}

		}

	}

	/*
	 * ??????????????
	 */
	private void PartsinformationProcessing(XSSFWorkbook book, List assynos, List assynamelist) {
		// TODO Auto-generated method stub
		// ??????????sheet
//		if (!updateflag) 
//		{
//			for(int i=0;i<deletelist.size();i++)
//			{
//				if(deletelist.get(i).toString().contains("??????"))
//				{
//					return;
//				}
//			}			
//		}	
		int sheetnum = 0;
		sheetnum = book.getNumberOfSheets();
		int sheetAtIndex = -1; // ??????????????
		for (int i = 0; i < sheetnum; i++) {
			String sheetname = book.getSheetName(i);
			if (sheetname.contains("??????")) {
				sheetAtIndex = i;
				break;
			}
		}
		if (sheetAtIndex == -1) {
			return;
		}
		// ????????????????????????
		int sum = 0;
//		for (Map.Entry<String, String> entry : fymap.entrySet()) {
//			sum = sum + Integer.parseInt(entry.getValue()) + 1;
//		}
		sum = partlist.size();
		// ??24????????sheet??
		int page = sum / 24 + 1;

		// ??????????????????????sheet????????????????
		if (sum % 24 == 0) {
			if (page > 1) {
				page = page - 1;
			}
		}

		int index = sheetAtIndex + 1;

		// ????????????
		Font font = book.createFont();
		font.setColor((short) 12);// ????????
		font.setFontName("MS PGothic");
		font.setFontHeightInPoints((short) 12);
		XSSFCellStyle style = book.createCellStyle();
		style.setBorderBottom(XSSFCellStyle.BORDER_THIN); // ??????
		style.setBorderLeft(XSSFCellStyle.BORDER_THIN);// ??????
		style.setBorderTop(XSSFCellStyle.BORDER_THIN);// ??????
		style.setBorderRight(XSSFCellStyle.BORDER_THIN);// ??????
		style.setVerticalAlignment(XSSFCellStyle.VERTICAL_CENTER);
		style.setAlignment(XSSFCellStyle.ALIGN_CENTER);
		style.setFont(font);

		Font font2 = book.createFont();
		font2.setFontName("MS PGothic");
		font2.setColor((short) 12);// ????????
		font2.setFontHeightInPoints((short) 12);
		font2.setBoldweight(Font.BOLDWEIGHT_BOLD); // ????????
		XSSFCellStyle style2 = book.createCellStyle();
		style2.setBorderBottom(XSSFCellStyle.BORDER_THIN); // ??????
		style2.setBorderLeft(XSSFCellStyle.BORDER_MEDIUM);// ??????
		style2.setBorderTop(XSSFCellStyle.BORDER_THIN);// ??????
		style2.setBorderRight(XSSFCellStyle.BORDER_THIN);// ??????
		style2.setVerticalAlignment(XSSFCellStyle.VERTICAL_CENTER);
		style2.setAlignment(XSSFCellStyle.ALIGN_CENTER);
		style2.setFont(font2);

		XSSFCellStyle style3 = book.createCellStyle();
		style3.setBorderBottom(XSSFCellStyle.BORDER_THIN); // ??????
		// style3.setBorderLeft(XSSFCellStyle.BORDER_THIN);// ??????
		style3.setBorderTop(XSSFCellStyle.BORDER_THIN);// ??????
		style3.setBorderRight(XSSFCellStyle.BORDER_THIN);// ??????
		style3.setVerticalAlignment(XSSFCellStyle.VERTICAL_CENTER);
		style3.setAlignment(XSSFCellStyle.ALIGN_CENTER);
		style3.setFont(font);

		XSSFCellStyle style4 = book.createCellStyle();
		style4.setBorderBottom(XSSFCellStyle.BORDER_THIN); // ??????
		style4.setBorderLeft(XSSFCellStyle.BORDER_THIN);// ??????
		style4.setBorderTop(XSSFCellStyle.BORDER_THIN);// ??????
		// style4.setBorderRight(XSSFCellStyle.BORDER_THIN);// ??????
		style4.setVerticalAlignment(XSSFCellStyle.VERTICAL_CENTER);
		style4.setAlignment(XSSFCellStyle.ALIGN_CENTER);
		style4.setFont(font);

		Font font3 = book.createFont();
		font3.setColor((short) 2);// ????????
		font3.setFontName("MS PGothic");
		font3.setFontHeightInPoints((short) 12);
		XSSFCellStyle style5 = book.createCellStyle();
		style5.setBorderBottom(XSSFCellStyle.BORDER_THIN); // ??????
		style5.setBorderLeft(XSSFCellStyle.BORDER_THIN);// ??????
		style5.setBorderTop(XSSFCellStyle.BORDER_THIN);// ??????
		style5.setBorderRight(XSSFCellStyle.BORDER_THIN);// ??????
		style5.setVerticalAlignment(XSSFCellStyle.VERTICAL_CENTER);
		style5.setAlignment(XSSFCellStyle.ALIGN_CENTER);
		style5.setFont(font3);

		XSSFCellStyle style6 = book.createCellStyle();
		style6.setBorderBottom(XSSFCellStyle.BORDER_THIN); // ??????
		// style3.setBorderLeft(XSSFCellStyle.BORDER_THIN);// ??????
		style6.setBorderTop(XSSFCellStyle.BORDER_THIN);// ??????
		style6.setBorderRight(XSSFCellStyle.BORDER_THIN);// ??????
		style6.setVerticalAlignment(XSSFCellStyle.VERTICAL_CENTER);
		style6.setAlignment(XSSFCellStyle.ALIGN_CENTER);
		style6.setFont(font3);

		XSSFCellStyle style7 = book.createCellStyle();
		style7.setBorderBottom(XSSFCellStyle.BORDER_THIN); // ??????
		style7.setBorderLeft(XSSFCellStyle.BORDER_THIN);// ??????
		style7.setBorderTop(XSSFCellStyle.BORDER_THIN);// ??????
		// style4.setBorderRight(XSSFCellStyle.BORDER_THIN);// ??????
		style7.setVerticalAlignment(XSSFCellStyle.VERTICAL_CENTER);
		style7.setAlignment(XSSFCellStyle.ALIGN_CENTER);
		style7.setFont(font3);

		/***********************************************/
		// ????????????????????????????????????????????????
		// ????????????????????????????????????????????????????????????????????????????
		List oldassynos = new ArrayList();
		List oldpartList = new ArrayList();
		List oldpartname = new ArrayList();
		if (updateflag) {
			int gcnum = 0;
			for (int i = 0; i < sheetnum; i++) {
				String sheetname = book.getSheetName(i);
				if (sheetname.contains("??????")) {
					gcnum++;
				}
			}
			// ????sheet????????????????????????????
			index = sheetAtIndex + page;

			// ??????????sheet????????????????????????????????????
			XSSFCell cell;
			XSSFRow row;
			for (int i = sheetAtIndex; i < sheetAtIndex + gcnum; i++) {
				XSSFSheet sheet = book.getSheetAt(i);
				// ????Assy List????
				for (int j = 0; j < 10; j++) {
					row = sheet.getRow(9 + j);
					String preassy;// ????????????????
					String suffixassy;// ????????????????
					if (row != null) {
						cell = row.getCell(7);
						if (cell != null) {
							preassy = convertCellValueToString(cell, Cell.CELL_TYPE_STRING);
							cell = row.getCell(13);
							if (cell != null) {
								suffixassy = convertCellValueToString(cell, Cell.CELL_TYPE_STRING);
								if ((preassy != null && !preassy.isEmpty())
										|| (suffixassy != null && !suffixassy.isEmpty())) {
									String allassy = preassy.trim() + suffixassy.trim();
									oldassynos.add(allassy);
								}
							}
						}
					}
					setStringCellAndStyle(sheet, "", 9 + j, 7, style, Cell.CELL_TYPE_STRING);// ????????????????
					setStringCellAndStyle(sheet, "", 9 + j, 13, style, Cell.CELL_TYPE_STRING);// ????????????????
					setStringCellAndStyle(sheet, "", 9 + j, 19, style3, Cell.CELL_TYPE_STRING);// ????????????
				}
				// ????Part List????
				for (int j = 0; j < 24; j++) {
					row = sheet.getRow(23 + j);
					String preassy;// ????????????????
					String suffixassy;// ????????????????
					if (row != null) {
						cell = row.getCell(13);
						if (cell != null) {
							preassy = convertCellValueToString(cell, Cell.CELL_TYPE_STRING);
							cell = row.getCell(18);
							if (cell != null) {
								suffixassy = convertCellValueToString(cell, Cell.CELL_TYPE_STRING);
								if ((preassy != null && !preassy.isEmpty())
										|| (suffixassy != null && !suffixassy.isEmpty())) {
									String allassy = preassy.trim() + suffixassy.trim();
									oldpartList.add(allassy);
								}
							}
							cell = row.getCell(24);
							if (cell != null) {
								String partname = convertCellValueToString(cell, Cell.CELL_TYPE_STRING);
								if (partname != null && !partname.isEmpty()) {
									oldpartname.add(partname);
								}
							}
						}
					}
					setStringCellAndStyle(sheet, "", 23 + j, 4, style2, Cell.CELL_TYPE_STRING);// ????
					setStringCellAndStyle(sheet, "", 23 + j, 7, style, Cell.CELL_TYPE_STRING);// ????????
					setStringCellAndStyle(sheet, "", 23 + j, 13, style4, Cell.CELL_TYPE_STRING);// ????????????
					setStringCellAndStyle(sheet, "", 23 + j, 18, style3, Cell.CELL_TYPE_STRING);// ????????????
					setStringCellAndStyle(sheet, "", 23 + j, 24, style3, Cell.CELL_TYPE_STRING);// ????????
					setStringCellAndStyle(sheet, "", 23 + j, 50, style, Cell.CELL_TYPE_STRING);// ????
					setStringCellAndStyle(sheet, "", 23 + j, 53, style, Cell.CELL_TYPE_STRING);// ????
					setStringCellAndStyle(sheet, "", 23 + j, 58, style, Cell.CELL_TYPE_STRING);// ????
					setStringCellAndStyle(sheet, "", 23 + j, 72, style, Cell.CELL_TYPE_STRING);// ????????
				}
			}
			if (gcnum < page) {
				for (int i = 1; i < page - gcnum; i++) {
					XSSFSheet newsheet = book.cloneSheet(sheetAtIndex);
					book.setSheetOrder(newsheet.getSheetName(), index);
					index++;
				}
			}

		} else {
			// ????page????1????????????sheet??
			if (page > 1) {
				for (int i = 1; i < page; i++) {
					XSSFSheet newsheet = book.cloneSheet(sheetAtIndex);
					book.setSheetOrder(newsheet.getSheetName(), index);
					index++;
				}
			}
		}
		/***********************************************/

		// ????????????
		int shnum = 0;
		for (int i = sheetAtIndex; i < index; i++) {
			int startrow = 23;// ??????????
			int endrow = 0;// ??????????
//			int Totalnum = 0;// ????????
			boolean flag = false;

			XSSFSheet sheet = book.getSheetAt(i);
			// ??partlist??????????
			if (assynos != null) {
				for (int k = 0; k < assynamelist.size(); k++) {
					String prename = "";
					String aftername = "";
					String[] assyVal = (String[]) assynamelist.get(k);
					String assyvalue = assyVal[0];
					System.out.println("????????????????????????" + assyvalue);
					if (assyvalue != null && assyvalue.length() > 5) {
						prename = assyvalue.substring(0, 5);
						aftername = assyvalue.substring(5).trim();
					} else {
						prename = assyvalue;
						aftername = "";
					}
					
					// ????????????????????????????????????????????????
					if (oldassynos != null && updateflag) {
						String allassy = prename.trim() + aftername.trim();
						if (!oldassynos.contains(allassy)) {
							XSSFCellStyle newstyle = getXSSFStyle(book, sheet, 9 + k, 7, 2, -1);
							XSSFCellStyle newstyle2 = getXSSFStyle(book, sheet, 9 + k, 13, 2, -1);
							setStringCellAndStyle2(sheet, prename, 9 + k, 7, newstyle, Cell.CELL_TYPE_STRING);// ????????????????
							setStringCellAndStyle2(sheet, aftername, 9 + k, 13, newstyle2, Cell.CELL_TYPE_STRING);// ????????????????
						} else {
							XSSFCellStyle newstyle = getXSSFStyle(book, sheet, 9 + k, 7, 12, -1);
							XSSFCellStyle newstyle2 = getXSSFStyle(book, sheet, 9 + k, 13, 12, -1);
							setStringCellAndStyle2(sheet, prename, 9 + k, 7, newstyle, Cell.CELL_TYPE_STRING);// ????????????????
							setStringCellAndStyle2(sheet, aftername, 9 + k, 13, newstyle2, Cell.CELL_TYPE_STRING);// ????????????????
						}
					} else {
						XSSFCellStyle newstyle = getXSSFStyle(book, sheet, 9 + k, 7, 12, -1);
						XSSFCellStyle newstyle2 = getXSSFStyle(book, sheet, 9 + k, 13, 12, -1);
						setStringCellAndStyle2(sheet, prename, 9 + k, 7, newstyle, Cell.CELL_TYPE_STRING);// ????????????????
						setStringCellAndStyle2(sheet, aftername, 9 + k, 13, newstyle2, Cell.CELL_TYPE_STRING);// ????????????????
					}
					setStringCellAndStyle(sheet, assyVal[1], 9 + k, 19, style3, Cell.CELL_TYPE_STRING);// ????????????
				}
			}
			// ??partlist????
			if (i == index - 1) {
				for (int j = 0; j + 24 * shnum < partlist.size(); j++) {
					String[] str = (String[]) partlist.get(j + 24 * shnum);
					// ??????????????
					if (str[7] != null) {
						String prename = "";
						String aftername = "";
						System.out.println("????????????????" + str[1]);
						//??	????????????????????????????????????5??????5??????PartList  20201118
						if(str[1] != null && (str[1].contains(" ") || str[1].contains(" ")))
						{
							if (str[1] != null && str[1].length() > 5) {
								prename = str[1].substring(0, 5);
								aftername = str[1].substring(5).trim();
							} else {
								prename = str[1];
								aftername = "";
							}
						}
						else
						{
							prename = "";
							aftername = str[1];
						}
						// ????????????????????????????????????????????????
						if (oldpartList != null && updateflag) {
							String allassyno = prename.trim() + aftername.trim();
							if (!oldpartList.contains(allassyno)) {
								XSSFCellStyle newstyle = getXSSFStyle(book, sheet, 23 + j, 13, 2, -1);
								XSSFCellStyle newstyle2 = getXSSFStyle(book, sheet, 23 + j, 18, 2, -1);
								setStringCellAndStyle2(sheet, prename, 23 + j, 13, newstyle, Cell.CELL_TYPE_STRING);// ????????????
								setStringCellAndStyle2(sheet, aftername, 23 + j, 18, newstyle2, Cell.CELL_TYPE_STRING);// ????????????
							} else {
								XSSFCellStyle newstyle = getXSSFStyle(book, sheet, 23 + j, 13, 12, -1);
								XSSFCellStyle newstyle2 = getXSSFStyle(book, sheet, 23 + j, 18, 12, -1);
								setStringCellAndStyle2(sheet, prename, 23 + j, 13, newstyle, Cell.CELL_TYPE_STRING);// ????????????
								setStringCellAndStyle2(sheet, aftername, 23 + j, 18, newstyle2, Cell.CELL_TYPE_STRING);// ????????????
							}
						} else {
							XSSFCellStyle newstyle = getXSSFStyle(book, sheet, 23 + j, 13, 12, -1);
							XSSFCellStyle newstyle2 = getXSSFStyle(book, sheet, 23 + j, 18, 12, -1);
							setStringCellAndStyle2(sheet, prename, 23 + j, 13, newstyle, Cell.CELL_TYPE_STRING);// ????????????
							setStringCellAndStyle2(sheet, aftername, 23 + j, 18, newstyle2, Cell.CELL_TYPE_STRING);// ????????????

						}
						if (oldpartname != null && updateflag) {
							if (!oldpartname.contains(str[2])) {
								XSSFCellStyle newstyle = getXSSFStyle(book, sheet, 23 + j, 24, 2, -1);
								setStringCellAndStyle2(sheet, str[2], 23 + j, 24, newstyle, Cell.CELL_TYPE_STRING);// ????????
							} else {
								XSSFCellStyle newstyle = getXSSFStyle(book, sheet, 23 + j, 24, 12, -1);
								setStringCellAndStyle2(sheet, str[2], 23 + j, 24, newstyle, Cell.CELL_TYPE_STRING);// ????????
							}

						} else {
							XSSFCellStyle newstyle = getXSSFStyle(book, sheet, 23 + j, 24, 12, -1);
							setStringCellAndStyle2(sheet, str[2], 23 + j, 24, newstyle, Cell.CELL_TYPE_STRING);// ????????
						}
						setStringCellAndStyle(sheet, str[7], 23 + j, 4, style2, Cell.CELL_TYPE_STRING);// ????
						setStringCellAndStyle(sheet, str[0], 23 + j, 7, style, Cell.CELL_TYPE_STRING);// ????????
						setStringCellAndStyle(sheet, str[3], 23 + j, 50, style, Cell.CELL_TYPE_STRING);// ????
						setStringCellAndStyle(sheet, str[4], 23 + j, 53, style, Cell.CELL_TYPE_STRING);// ????
						setStringCellAndStyle(sheet, str[5], 23 + j, 58, style, Cell.CELL_TYPE_STRING);// ????
						setStringCellAndStyle(sheet, str[6], 23 + j, 72, style, Cell.CELL_TYPE_STRING);// ????????

						flag = true;
						endrow = 23 + j;
//						if (str[3] != null && !str[3].isEmpty()) {
//							Totalnum = Totalnum + Integer.parseInt(str[3]);
//						} else {
//							Totalnum = Totalnum + 0;
//						}

					} else {
					}

				}
			} else {
				for (int j = 0; j + 24 * shnum < 24 + 24 * shnum; j++) {
					// ??????????????
					String[] str = (String[]) partlist.get(j + 24 * shnum);
					// ??????????????
					if (str[7] != null) {
						String prename = "";
						String aftername = "";
						//??	????????????????????????????????????5??????5??????PartList  20201118
						if(str[1] != null && (str[1].contains(" ") || str[1].contains(" ")))
						{
							if (str[1] != null && str[1].length() > 5) {
								prename = str[1].substring(0, 5);
								aftername = str[1].substring(5).trim();
							} else {
								prename = str[1];
								aftername = "";
							}
						}
						else
						{
							prename = "";
							aftername = str[1];
						}
						// ????????????????????????????????????????????????
						if (oldpartList != null && updateflag) {
							String allassyno = prename.trim() + aftername.trim();
							if (!oldpartList.contains(allassyno)) {
								XSSFCellStyle newstyle = getXSSFStyle(book, sheet, 23 + j, 13, 2, -1);
								XSSFCellStyle newstyle2 = getXSSFStyle(book, sheet, 23 + j, 18, 2, -1);
								setStringCellAndStyle2(sheet, prename, 23 + j, 13, newstyle, Cell.CELL_TYPE_STRING);// ????????????
								setStringCellAndStyle2(sheet, aftername, 23 + j, 18, newstyle2, Cell.CELL_TYPE_STRING);// ????????????
							} else {
								XSSFCellStyle newstyle = getXSSFStyle(book, sheet, 23 + j, 13, 12, -1);
								XSSFCellStyle newstyle2 = getXSSFStyle(book, sheet, 23 + j, 18, 12, -1);
								setStringCellAndStyle2(sheet, prename, 23 + j, 13, newstyle, Cell.CELL_TYPE_STRING);// ????????????
								setStringCellAndStyle2(sheet, aftername, 23 + j, 18, newstyle2, Cell.CELL_TYPE_STRING);// ????????????
							}
						} else {
							XSSFCellStyle newstyle = getXSSFStyle(book, sheet, 23 + j, 13, 12, -1);
							XSSFCellStyle newstyle2 = getXSSFStyle(book, sheet, 23 + j, 18, 12, -1);
							setStringCellAndStyle2(sheet, prename, 23 + j, 13, newstyle, Cell.CELL_TYPE_STRING);// ????????????
							setStringCellAndStyle2(sheet, aftername, 23 + j, 18, newstyle2, Cell.CELL_TYPE_STRING);// ????????????

						}
						if (oldpartname != null && updateflag) {
							if (!oldpartname.contains(str[2])) {
								XSSFCellStyle newstyle = getXSSFStyle(book, sheet, 23 + j, 24, 2, -1);
								setStringCellAndStyle2(sheet, str[2], 23 + j, 24, newstyle, Cell.CELL_TYPE_STRING);// ????????
							} else {
								XSSFCellStyle newstyle = getXSSFStyle(book, sheet, 23 + j, 24, 12, -1);
								setStringCellAndStyle2(sheet, str[2], 23 + j, 24, newstyle, Cell.CELL_TYPE_STRING);// ????????
							}

						} else {
							XSSFCellStyle newstyle = getXSSFStyle(book, sheet, 23 + j, 24, 12, -1);
							setStringCellAndStyle2(sheet, str[2], 23 + j, 24, newstyle, Cell.CELL_TYPE_STRING);// ????????
						}
						setStringCellAndStyle(sheet, str[7], 23 + j, 4, style2, Cell.CELL_TYPE_STRING);// ????
						setStringCellAndStyle(sheet, str[0], 23 + j, 7, style, Cell.CELL_TYPE_STRING);// ????????
						setStringCellAndStyle(sheet, str[3], 23 + j, 50, style, Cell.CELL_TYPE_STRING);// ????
						setStringCellAndStyle(sheet, str[4], 23 + j, 53, style, Cell.CELL_TYPE_STRING);// ????
						setStringCellAndStyle(sheet, str[5], 23 + j, 58, style, Cell.CELL_TYPE_STRING);// ????
						setStringCellAndStyle(sheet, str[6], 23 + j, 72, style, Cell.CELL_TYPE_STRING);// ????????

						flag = true;
						endrow = 23 + j;
//						if (str[3] != null && !str[3].isEmpty()) {
//							Totalnum = Totalnum + Integer.parseInt(str[3]);
//						} else {
//							Totalnum = Totalnum + 0;
//						}

					} else {
					}
				}
			}
			shnum++;
		}

	}

	/*
	 * ????????????
	 */
	private void getPartsinformation(TCComponentBOMLine gwbl) throws TCException {
		// TODO Auto-generated method stub
//		ArrayList install = new ArrayList();
//		ArrayList templist = new ArrayList();
//		// ??????????????????????????????
//		install = Util.getChildrenByBOMLine(gwbl, "B8_BIWOperationRevision");
//
//		for (int i = 0; i < install.size(); i++) {
//			TCComponentBOMLine bl = (TCComponentBOMLine) install.get(i);
//			ArrayList bflist = new ArrayList();
//			bflist = Util.getChildrenByBOMLine(bl, "DFL9SolItmPartRevision");
//			for (int j = 0; j < bflist.size(); j++) {
//				String[] info = new String[8];
//				TCComponentBOMLine bfbl = (TCComponentBOMLine) bflist.get(j);
//				info[0] = Util.getProperty(bfbl, "bl_sequence_no");// ????????
//				info[1] = Util.getProperty(bfbl, "bl_DFL9SolItmPartRevision_dfl9_part_no");// ????????
//				// info[2] = Util.getProperty(bfbl, "bl_rev_object_name");// ????????
//				info[2] = Util.getProperty(bfbl.getItemRevision(), "dfl9_CADObjectName");// ????????
//				info[3] = Util.getProperty(bfbl, "bl_quantity");// ????
//				if (info[3] == null || info[3].isEmpty()) {
//					info[3] = "1";
//				}
//				String partresoles = "";
//				partresoles = Util.getProperty(bfbl, "B8_NoteManualMark");// ???????? ??????
//				if (partresoles == null || partresoles.isEmpty()) {
//					partresoles = Util.getProperty(bfbl, "B8_NoteIsBiwTrUnit");// ???????? ??????
//				}
//				if (partresoles.equals("??????")) {
//					partresoles = "????????";
//				}
//				info[6] = partresoles;
//
//				if (partresoles.equals("??????")) {
//					String bh = Util.getProperty(bfbl, "bl_DFL9SolItmPartRevision_dfl9PartThickness");
//					if (Util.isNumber(bh)) {
//						System.out.println("??????????" + bh);
//						info[4] = String.format("%.2f", Double.parseDouble(bh));// ????
//					} else {
//						info[4] = bh;// ????
//					}
//					info[5] = Util.getProperty(bfbl, "bl_DFL9SolItmPartRevision_dfl9PartMaterial");// ????
//				} else {
//					info[4] = "";// ????
//					info[5] = "";// ????
//				}
//				templist.add(info);
//			}
//		}
//		// ????????????????????????assy????
//		TCProperty pp = gwbl.getTCProperty("Mfg0predecessors");
//		if (pp != null) {
//			TCComponent[] obj = pp.getReferenceValueArray();
//			for (int i = 0; i < obj.length; i++) {
//				TCComponentBOMLine prebl = (TCComponentBOMLine) obj[i];
//				String sequence_no = Util.getProperty(prebl, "bl_sequence_no");// ????????
//				String quantity = Util.getProperty(prebl, "bl_quantity");// ????
//				if (quantity == null || quantity.isEmpty()) {
//					quantity = "1";
//				}
//				// ???????????? ,??????????????????????#??????????????????+??????????????????????????
//				String linename = Util.getProperty(prebl.parent(), "bl_rev_object_name");
//				String staname = Util.getProperty(prebl, "bl_rev_object_name");
//				String assyname = "";
//				if (staname.length() > 1) {
//					if (staname.substring(0, 1).equals("#")) {
//						assyname = linename + " " + staname;
//					} else {
//						assyname = staname;
//					}
//				} else {
//					assyname = linename + " " + staname;
//				}
//
//				TCProperty p = prebl.getItemRevision().getTCProperty("b8_ProcAssyNo2");
//				String[] assynos;
//				if (p != null) {
//					assynos = p.getStringValueArray();// ???????? ????
//				} else {
//					assynos = null;
//				}
//				if (assynos != null && assynos.length > 0) {
//					for (int j = 0; j < assynos.length; j++) {
//						String[] info = new String[8];
//						info[0] = sequence_no;// ????????
//						info[1] = assynos[j];// ????????
//						info[2] = assyname;// ????????
//						info[3] = quantity;// ????
//						info[4] = "";// ????
//						info[5] = "";// ????
//						info[6] = "????????";// ???????? ??????
//						templist.add(info);
//					}
//				}
//			}
//		}
//		// ????????????????????????????????????
//		Map<String, String[]> map = new HashMap<String, String[]>();
//		for (int i = 0; i < templist.size(); i++) {
//			String[] value = (String[]) templist.get(i);
//			String key = value[1];
//			if (!map.containsKey(key)) {
//				map.put(key, value);
//			} else {
//				String[] oldstr = map.get(key);
//				int quality = 0;
//				quality = Integer.parseInt(oldstr[3]) + Integer.parseInt(value[3]);
//				oldstr[3] = Integer.toString(quality);
//				map.put(key, oldstr);
//			}
//		}
//		ArrayList newtemplist = new ArrayList();
//		for (Map.Entry<String, String[]> entry : map.entrySet()) {
//			String[] values = entry.getValue();
//			newtemplist.add(values);
//		}
//
//		// ????????????????????????????
//		Comparator comparator = getComParatorBysequenceno();
//		Collections.sort(newtemplist, comparator);
//
//		int label = 0; // ????????
//		int num = 1;// ??????????????????????
//		String prePartno = "";// ??????????5??????
//		String[] bh = { "A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S",
//				"T", "U", "V", "W", "X", "Y", "Z" };
//		// ????????
//		Map<String, String> tempmap = new HashMap<String, String>();
//		ArrayList tempPartlist = new ArrayList();
//		for (int i = 0; i < newtemplist.size(); i++) {
//			String[] str = (String[]) newtemplist.get(i);
//			if (str[1].toString().length() > 5) {
//
//				prePartno = str[1].toString().substring(0, 5);
//			} else {
//				prePartno = str[1].toString();
//			}
//			String note = tempmap.get(prePartno);
//			// ??????????5??????????????????
//			if (note != null && !note.isEmpty()) {
//				str[7] = note;
//				String strnum = fymap.get(note);
//				int newnum = Integer.parseInt(strnum) + 1;
//				fymap.put(note, Integer.toString(newnum));
//			} else {
//				if (label < 26) {
//					str[7] = bh[label];
//				} else {
//					str[7] = "";
//					System.out.println("??????????????????????????");
//				}
//
//				fymap.put(bh[label], "1");
//				tempmap.put(prePartno, bh[label]);
//				label++;
//			}
//			tempPartlist.add(str);
//
//		}
//		// ????????????
//		Comparator comparator2 = getComParatorBybh();
//		Collections.sort(tempPartlist, comparator2);
//
//		String firstNo = "";
//		for (int i = 0; i < tempPartlist.size(); i++) {
//			String[] value = (String[]) tempPartlist.get(i);
//			if (i == 0) {
//				firstNo = value[7];
//				partlist.add(value);
//			} else {
//				if (!firstNo.equals(value[7].toString())) {
//					String[] str = new String[8];
//					partlist.add(str);
//					partlist.add(value);
//					firstNo = value[7];
//				} else {
//					partlist.add(value);
//				}
//			}
//		}
//		System.out.println(partlist);

	}

	private Comparator getComParatorBybh() {
		// TODO Auto-generated method stub
		Comparator comparator = new Comparator() {

			public int compare(Object obj, Object obj1) {
				// System.setProperty("java.util.Arrays.useLegacyMergeSort", "true");
				String[] comp1 = (String[]) obj;
				String[] comp2 = (String[]) obj1;

				String d1 = "";
				String d2 = "";
				if (obj != null && comp1[7] != null && !comp1[7].isEmpty()) {
					d1 = comp1[7].toString();
				}
				if (obj1 != null && comp2[7] != null && !comp2[7].isEmpty()) {
					d2 = comp2[7];
				}
				return d1.compareTo(d2);
			}
		};

		return comparator;
	}

	/*
	 * ????????????
	 */
	private void ValidPageProcessing(XSSFWorkbook book) {
		// TODO Auto-generated method stub
		XSSFSheet sheet = null;
		int sheetnum = 0;
		sheetnum = book.getNumberOfSheets();
		for (int i = 0; i < sheetnum; i++) {
			String sheetname = book.getSheetName(i);
			if (sheetname.contains("??????")) {
				sheet = book.getSheetAt(i);
				break;
			}
		}
		if (sheet == null) {
			return;
		}
		// ????????????????sheet??????????120??????????????????
		// ????????????
		Font font = book.createFont();
		font.setColor((short) 12);// ????????
		font.setFontName("????");
		font.setFontHeightInPoints((short) 14);
		XSSFCellStyle style = book.createCellStyle();
		style.setBorderBottom(XSSFCellStyle.BORDER_THIN); // ??????
		style.setBorderLeft(XSSFCellStyle.BORDER_THIN);// ??????
		style.setBorderTop(XSSFCellStyle.BORDER_THIN);// ??????
		style.setBorderRight(XSSFCellStyle.BORDER_THIN);// ??????
		style.setVerticalAlignment(XSSFCellStyle.VERTICAL_CENTER);
		style.setAlignment(XSSFCellStyle.ALIGN_CENTER);
		style.setFont(font);

		int page = (sheetnum - 1) / 40 + 1;

		for (int i = 0; i < page; i++) {
			if (i == page - 1) {
				for (int j = 0; j < sheetnum - 40 * i; j++) {
					setStringCellAndStyle2(sheet, "??", 7 + j, 11 + 35 * i, style, Cell.CELL_TYPE_STRING); // ????
				}
			} else {
				for (int j = 0; j < 40; j++) {
					setStringCellAndStyle2(sheet, "??", 7 + j, 11 + 35 * i, style, Cell.CELL_TYPE_STRING); // ????
				}
			}
		}
	}

	/*
	 * ????????sheet??????????
	 */
	private void SetSheetRename(XSSFWorkbook book) {
		// TODO Auto-generated method stub

		int sheetnum = book.getNumberOfSheets();
		// ????????????,??????????????sheet????????????????????????????????????????1,2......
		String tempname = "";
		String sheetAllname;
		int num = 1;
		Pattern p = Pattern.compile("[0-9a-fA-F]"); // ??????????
		Pattern p2 = Pattern.compile("[0-9]"); // ??????????
		// ????????????????????????????????????????????????????
		for (int i = 0; i < sheetnum; i++) {
			XSSFSheet sheet = book.getSheetAt(i);
			String sheetname = sheet.getSheetName();
			sheetAllname = sheetname + (i + 1);
			book.setSheetName(i, sheetAllname);
		}

		for (int i = 0; i < sheetnum; i++) {
			XSSFSheet sheet = book.getSheetAt(i);
			// ????????????????????
			String sheetname = "";
			String oldsheetname = sheet.getSheetName();
			if (oldsheetname.length() > 2) {
				Matcher m = p.matcher(oldsheetname.substring(0, 3));
				Matcher m2 = p2.matcher(oldsheetname.substring(3));
				sheetname = m.replaceAll("").trim() + m2.replaceAll("").trim();
			} else {
				sheetname = oldsheetname;
			}
			// String sheetname = sheet.getSheetName().substring(2);
			// ????????????????????????????????????????
			if (i == 0) {
				tempname = sheetname;
				sheetAllname = String.format("%02d", i + 1) + sheetname;
				book.setSheetName(i, sheetAllname);

			} else {
				if (sheetname.contains(tempname)) {
					// ????num??1????????sheet??????????????????????????????????
					if (num == 1) {
						sheetAllname = String.format("%02d", i) + tempname + num;
						book.setSheetName(i - 1, sheetAllname);
					}
					sheetAllname = String.format("%02d", i + 1) + tempname + (num + 1);
					book.setSheetName(i, sheetAllname);
					num++;
				} else {
					sheetAllname = String.format("%02d", i + 1) + sheetname;
					tempname = sheetname;
					book.setSheetName(i, sheetAllname);

					num = 1;
				}
			}
			// ????????????
			book.removePrintArea(i);
		}
		for (int i = 0; i < sheetnum; i++) {
			XSSFSheet sheet = book.getSheetAt(i);
			book.setPrintArea(i, 0, 114, 0, 51);
			PrintSetup printSetup = sheet.getPrintSetup();
			printSetup.setPaperSize(XSSFPrintSetup.A4_PAPERSIZE);
			printSetup.setScale((short) 70);// ????????????????100????????
			printSetup.setLandscape(true); // ??????????true????????false??????(????)

		}
	}

	/*
	 * ??????sheet????????????
	 */
	private void writePublicDataToSheet(XSSFWorkbook book, ArrayList plist) {
		// TODO Auto-generated method stub
		// ????????
		Font font = book.createFont();
		font.setColor((short) 12);
		font.setFontName("????");
		font.setFontHeightInPoints((short) 16);
		// ????????????
		XSSFCellStyle cellStyle1 = book.createCellStyle();
		cellStyle1.setBorderBottom(XSSFCellStyle.BORDER_THIN);
		cellStyle1.setBorderLeft(XSSFCellStyle.BORDER_THIN);
		cellStyle1.setBorderRight(XSSFCellStyle.BORDER_THIN);
		cellStyle1.setBorderTop(XSSFCellStyle.BORDER_THIN);
		cellStyle1.setVerticalAlignment(XSSFCellStyle.VERTICAL_CENTER);// ????????????
		cellStyle1.setAlignment(XSSFCellStyle.ALIGN_CENTER);
		cellStyle1.setFont(font);

		XSSFCellStyle cellStyle2 = book.createCellStyle();
		Font font2 = book.createFont();
		font2.setColor(IndexedColors.BLUE.getIndex());
		font2.setBoldweight(Font.BOLDWEIGHT_BOLD); // ????????
		font2.setFontHeightInPoints((short) 16);
		font2.setFontName("????");
		cellStyle2.setVerticalAlignment(XSSFCellStyle.VERTICAL_CENTER);// ????????????
		cellStyle2.setAlignment(XSSFCellStyle.ALIGN_LEFT);
		cellStyle2.setFont(font2);

		XSSFCellStyle cellStyle3 = book.createCellStyle();
		Font font3 = book.createFont();
		font3.setColor(IndexedColors.BLUE.getIndex());
		// font3.setBoldweight(Font.BOLDWEIGHT_BOLD); // ????????
		font3.setItalic(true); // ??????????
		font3.setFontHeightInPoints((short) 72);
		font3.setFontName("????");
		cellStyle3.setVerticalAlignment(XSSFCellStyle.VERTICAL_CENTER);// ????????????
		cellStyle3.setAlignment(XSSFCellStyle.ALIGN_CENTER);
		cellStyle3.setFont(font3);

		// ????????sheet??????????????????????
		int sheetnum = book.getNumberOfSheets();
		for (int n = 0; n < sheetnum; n++) {
			XSSFSheet sh = book.getSheetAt(n);
			String sheetname = sh.getSheetName();
			if (sheetname.contains("????")) {

				/**************************************************/
				// ????????????????????????????????????????????????
//				if (updateflag) {
//					List<String> delPicturesList = ReportUtils.removePictrues07((XSSFSheet) sh, (XSSFWorkbook) book, 3,
//							48, 100, 115);
//					System.out.println("-----------????????????????-----------");
//					for (String name : delPicturesList) {
//						System.out.println(name);
//					}
//				}
				/**************************************************/
				if (!updateflag && !model.equals("??????????")) { // ????????????????????????????????
					setStringCellAndStyle(sh, plist.get(3).toString(), 22, 5, cellStyle3, Cell.CELL_TYPE_STRING); // ??????????????????
				}
				if (Import != null && Import.size() > 0) {
					for (int m = 0; m < Import.size(); m++) {
						InputStream is = null;
						if (Import.get(m).toString().trim().equals("A")) {
							is = this.getClass().getResourceAsStream("/com/dfl/report/imags/A.png");
						}
						if (Import.get(m).toString().trim().equals("B")) {
							is = this.getClass().getResourceAsStream("/com/dfl/report/imags/B.png");
						}
						if (is != null) {
							if (!updateflag) {
								writepicturetosheet(book, sh, is, 105, 5 + m * 5, 111, 9 + m * 5);
							}
						}
					}
				}
			}
			// ??????
			setStringCellAndStyle(sh, plist.get(6).toString(), 2, 0, cellStyle1, Cell.CELL_TYPE_STRING); // ????
			// ?????????? ??????????????????????????????????
			if (!updateflag) {
				setStringCellAndStyle(sh, plist.get(0).toString(), 2, 6, cellStyle1, Cell.CELL_TYPE_STRING); // ????
				setStringCellAndStyle(sh, plist.get(1).toString(), 2, 30, cellStyle1, Cell.CELL_TYPE_STRING);// ????
				setStringCellAndStyle(sh, plist.get(5).toString(), 48, 108, cellStyle2, Cell.CELL_TYPE_STRING);// ????
			}
			setStringCellAndStyle(sh, plist.get(2).toString(), 2, 90, cellStyle1, Cell.CELL_TYPE_STRING);// ????
			if (!updateflag && !model.equals("??????????")) { // ????????????????????????????????
				setStringCellAndStyle(sh, plist.get(3).toString(), 50, 72, cellStyle2, Cell.CELL_TYPE_STRING);// ????????
				setStringCellAndStyle(sh, plist.get(4).toString(), 51, 94, cellStyle2, Cell.CELL_TYPE_STRING);// ????????
			}
			setStringCellAndStyle(sh, Integer.toString(n + 1), 50, 107, cellStyle2, 10);// ????????
			setStringCellAndStyle(sh, Integer.toString(sheetnum), 50, 112, cellStyle2, 10);// ??????

			// ???????? ??????????
//			XSSFRow row = sh.getRow(50);
//			if (row != null) {
//				XSSFCell cell = row.getCell(72);
//				if (cell != null) {
//					NewOutputDataToExcel.setFontSize(book, cell, (short) 16);
//				}
//			}
		}
	}

	/*
	 * ??????sheet????????????
	 */
	private void writeRepatPublicDataToSheet(XSSFWorkbook book) {
		// TODO Auto-generated method stub
		// ????????
		Font font = book.createFont();
		font.setColor((short) 12);
		font.setFontName("????");
		font.setFontHeightInPoints((short) 16);
		// ????????????
		XSSFCellStyle cellStyle1 = book.createCellStyle();
		cellStyle1.setBorderBottom(XSSFCellStyle.BORDER_THIN);
		cellStyle1.setBorderLeft(XSSFCellStyle.BORDER_THIN);
		cellStyle1.setBorderRight(XSSFCellStyle.BORDER_THIN);
		cellStyle1.setBorderTop(XSSFCellStyle.BORDER_THIN);
		cellStyle1.setVerticalAlignment(XSSFCellStyle.VERTICAL_CENTER);// ????????????
		cellStyle1.setAlignment(XSSFCellStyle.ALIGN_CENTER);
		cellStyle1.setFont(font);

		XSSFCellStyle cellStyle2 = book.createCellStyle();
		Font font2 = book.createFont();
		font2.setColor(IndexedColors.BLUE.getIndex());
		font2.setBoldweight(Font.BOLDWEIGHT_BOLD); // ????????
		font2.setFontHeightInPoints((short) 16);
		font2.setFontName("????");
		cellStyle2.setVerticalAlignment(XSSFCellStyle.VERTICAL_CENTER);// ????????????
		cellStyle2.setAlignment(XSSFCellStyle.ALIGN_LEFT);
		cellStyle2.setFont(font2);

		// ????????sheet??????????????????????
		int sheetnum = book.getNumberOfSheets();
		for (int n = 0; n < sheetnum; n++) {
			XSSFSheet sh = book.getSheetAt(n);
			String sheetname = sh.getSheetName();
			setStringCellAndStyle(sh, Integer.toString(n + 1), 50, 107, cellStyle2, 10);// ????????
			setStringCellAndStyle(sh, Integer.toString(sheetnum), 50, 112, cellStyle2, 10);// ??????

		}
	}

	/*
	 * ??????????????sheet????????????????
	 */
	private XSSFWorkbook creatEngineeringXSSFWorkbook(InputStream inputStream, ArrayList list,
			LinkedHashMap<String, String> map) {
		// TODO Auto-generated method stub
		XSSFWorkbook book = null;
		try {
			book = new XSSFWorkbook(inputStream);

			// ????????sheet????????????????????????
			int sheetnum = book.getNumberOfSheets();
			deletelist = new ArrayList();
			ArrayList copylist = new ArrayList();
			Map<String, Integer> pxmap = new LinkedHashMap<String, Integer>();// ????sheet????
			for (int i = 0; i < sheetnum; i++) {
				String allshname = book.getSheetName(i);
				String sheetname = allshname.substring(2);
				if (!list.contains(sheetname)) {
					deletelist.add(book.getSheetName(i));
				} else {
					int shat = 0;
					for (int j = 0; j < list.size(); j++) {
						String strval = (String) list.get(j);
						if (strval.equals(sheetname)) {
							shat = j;
							break;
						}
					}
					pxmap.put(allshname, shat);
					// ????????????????sheet??????????sheet??
					int sheetNum = Integer.parseInt(map.get(sheetname));
					// ????????????1????????????sheet??
					if (sheetNum > 1) {
						copylist.add(book.getSheetName(i));
					}
				}
			}
			for (Map.Entry<String, Integer> entry : pxmap.entrySet()) {
				String key = entry.getKey();
				Integer value = entry.getValue();
				book.setSheetOrder(key, value);
			}

			// ??????????????sheet
			for (int k = 0; k < copylist.size(); k++) {
				String sheetAllname = (String) copylist.get(k);
				int sheetNums = Integer.parseInt(map.get(sheetAllname.substring(2)));
				int sheetAt = book.getSheetIndex(sheetAllname);
				int index = sheetAt + 1;
				for (int n = 1; n < sheetNums; n++) {
					XSSFSheet newsheet = book.cloneSheet(sheetAt);
					book.setSheetOrder(newsheet.getSheetName(), index);
					index++;
				}
			}

		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		return book;
	}

	/*
	 * ??????????sheet??
	 */
	private void deleteSheets(XSSFWorkbook book) {
		if (deletelist != null && deletelist.size() > 0) {
			for (int j = 0; j < deletelist.size(); j++) {
//				System.out.println("sheet??????" + deletelist.get(j).toString() + " " + book.getSheetIndex(deletelist.get(j).toString()));
				book.removeSheetAt(book.getSheetIndex(deletelist.get(j).toString()));
			}
			List delist = new ArrayList();
			for (int i = 0; i < book.getNumberOfSheets(); i++) {
				String sheetname = book.getSheetName(i);
				boolean flag = getIsContain(sheetname);
				if (flag) {
					delist.add(sheetname);
				}
			}
			if (delist != null && delist.size() > 0) {
				for (int k = 0; k < delist.size(); k++) {
					book.removeSheetAt(book.getSheetIndex(delist.get(k).toString()));
				}
			}
		}
	}

	private boolean getIsContain(String name) {

		if (deletelist != null && deletelist.size() > 0) {
			for (int i = 0; i < deletelist.size(); i++) {
				String shname = (String) deletelist.get(i);
				if (name.contains(shname)) {
					return true;
				}
			}
		}
		return false;
	}

	// ????????????
	public static void setStringCellAndStyle(XSSFSheet sheet, String value, int rowIndex, int cellIndex,
			XSSFCellStyle Style, int celltype) {

		// ?????????????????????? 10????????11??double??

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

	public static void setStringCellAndStyle2(XSSFSheet sheet, String value, int rowIndex, int cellIndex,
			XSSFCellStyle Style, int celltype) {

		// ?????????????????????? 10????????11??double??

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
        if(Style!=null)
        {
        	cell.setCellStyle(Style);
        }		

	}

	private Comparator getComParatorBysequenceno() {
		// TODO Auto-generated method stub
		Comparator comparator = new Comparator() {

			public int compare(Object obj, Object obj1) {
				Object[] comp1 = (Object[]) obj;
				Object[] comp2 = (Object[]) obj1;

				Double d1 = 0.0;
				Double d2 = 0.0;
				if (comp1[0] != null && !comp1[0].toString().isEmpty()) {
					d1 = Double.parseDouble(comp1[0].toString());
				}
				if (comp2[0] != null && !comp2[0].toString().isEmpty()) {
					d2 = Double.parseDouble(comp2[0].toString());
				}
				if (d2 > d1) {
					return -1;
				}
				if (d2 == d1) {
					return 0;
				}

				return 1;
			}
		};

		return comparator;
	}

	/*
	 * ??????????????????????
	 */
	private Comparator getComParatorBylayersnum() {
		// TODO Auto-generated method stub
		Comparator comparator = new Comparator() {

			public int compare(Object obj, Object obj1) {
				// System.setProperty("java.util.Arrays.useLegacyMergeSort", "true");
				WeldPointBoardInformation comp1 = (WeldPointBoardInformation) obj;
				WeldPointBoardInformation comp2 = (WeldPointBoardInformation) obj1;
//				if (comp1.equals(comp2)) {
//					//System.out.println("*************************");
//					return 0;
//				}

				int d1 = 0;
				int d2 = 0;
				if (obj != null && comp1.getLayersnum() != null && !comp1.getLayersnum().isEmpty()) {
					d1 = Integer.parseInt(comp1.getLayersnum());
				}
				if (obj1 != null && comp2.getLayersnum() != null && !comp2.getLayersnum().isEmpty()) {
					d2 = Integer.parseInt(comp2.getLayersnum());
				}
				if (d2 < d1) {
					return -1;
				}
				if (d2 == d1) {
					return 0;
				}

				return 1;
			}
		};

		return comparator;
	}

	/**
	 * ???????? Region
	 * 
	 * @param sheet
	 * @param row
	 * @param column
	 * @return
	 */
	public static int getMergedRegionIndex(XSSFSheet sheet, int row, int column) {
		int sheetMergeCount = sheet.getNumMergedRegions();

		for (int i = 0; i < sheetMergeCount; i++) {
			CellRangeAddress ca = sheet.getMergedRegion(i);
			int firstColumn = ca.getFirstColumn();
			int lastColumn = ca.getLastColumn();
			int firstRow = ca.getFirstRow();
			int lastRow = ca.getLastRow();

			if (row >= firstRow && row <= lastRow) {
				if (column >= firstColumn && column <= lastColumn) {
					return i;
				}
			}
		}

		return -1;
	}

	/*
	 * ????????
	 */
	private String getMinnum(String str1, String str2, String str3) {
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
		if (minstr.equals("9999")) {
			minstr = "";
		}
		return minstr;
	}

	/*
	 * ????????????????1180??????????????
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

	/*
	 * ??????????1.2g????????????????????/??????????????????????????true????????????false
	 */
	private boolean getColorDistinction(String layersnum, String partmaterial1, String partmaterial2,
			String partmaterial3, String partthickness1, String partthickness2, String partthickness3) {
		boolean flag = false;
		if (layersnum != null && !layersnum.isEmpty()) {
			int bznum = Integer.parseInt(layersnum);// ??????
			if (bznum == 1) {
				flag = true;
			} else if (bznum == 2) { // ??????????
				// ????????????1.2g??????
				boolean flag1 = getIscontains1180(partmaterial1);
				boolean flag2 = getIscontains1180(partmaterial2);
				boolean flag3 = getIscontains1180(partmaterial3);
				// ??????????????????????????????
				if (partmaterial1 == null || partmaterial1.isEmpty()) {
					// ????????1.2g??????
					if (flag2 && flag3) {
						flag = false;
					} else {
						flag = getCompareresultByTwo(partmaterial2, partmaterial3, partthickness2, partthickness3,
								flag2, flag3);
					}
				} else if (partmaterial2 == null || partmaterial2.isEmpty()) {
					// ????????1.2g??????
					if (flag1 && flag3) {
						flag = false;
					} else {
						flag = getCompareresultByTwo(partmaterial1, partmaterial3, partthickness1, partthickness3,
								flag1, flag3);
					}
				} else {
					// ????????1.2g??????
					if (flag1 && flag2) {
						flag = false;
					} else {
						flag = getCompareresultByTwo(partmaterial1, partmaterial2, partthickness1, partthickness2,
								flag1, flag2);
					}
				}
			} else { // ??????????
				// ????????????1.2g??????
				boolean flag1 = getIscontains1180(partmaterial1);
				boolean flag2 = getIscontains1180(partmaterial2);
				boolean flag3 = getIscontains1180(partmaterial3);
				// ??????????????1.2G ??????
				if (flag1 && flag2 && flag3) {
					flag = false;
				} else if (!flag1 && flag2 && flag3) { // ????1????1.2g??????????1.2g????
					// ????1.2g????????
					double h2 = getDoubleByString(partthickness2);
					double h3 = getDoubleByString(partthickness3);
					if (h2 < h3) {
						flag = getCompareresultByTwo(partmaterial1, partmaterial2, partthickness1, partthickness2,
								flag1, flag2);
					} else {
						flag = getCompareresultByTwo(partmaterial1, partmaterial3, partthickness1, partthickness3,
								flag1, flag3);
					}

				} else if (flag1 && !flag2 && flag3) { // ????2????1.2g??????????1.2g????
					// ????1.2g????????
					double h1 = getDoubleByString(partthickness1);
					double h3 = getDoubleByString(partthickness3);
					if (h1 < h3) {
						flag = getCompareresultByTwo(partmaterial2, partmaterial1, partthickness2, partthickness1,
								flag2, flag1);
					} else {
						flag = getCompareresultByTwo(partmaterial2, partmaterial3, partthickness2, partthickness3,
								flag2, flag3);
					}
				} else if (flag1 && flag2 && !flag3) { // ????3????1.2g??????????1.2g????
					// ????1.2g????????
					double h1 = getDoubleByString(partthickness1);
					double h2 = getDoubleByString(partthickness2);
					if (h1 < h2) {
						flag = getCompareresultByTwo(partmaterial3, partmaterial1, partthickness3, partthickness1,
								flag3, flag1);
					} else {
						flag = getCompareresultByTwo(partmaterial3, partmaterial2, partthickness3, partthickness2,
								flag3, flag2);
					}
				} else {// ??????????1.2g??????
					double h1 = getDoubleByString(partthickness1);
					double h2 = getDoubleByString(partthickness2);
					double h3 = getDoubleByString(partthickness3);
					int kn1 = getSheetstrength(partmaterial1);
					int kn2 = getSheetstrength(partmaterial2);
					int kn3 = getSheetstrength(partmaterial3);

					if (h1 != h2 && h1 != h3 && h2 != h3) { // ??????????
						// 1.2G ????????????????????????????
						if (flag1) {
							if (h1 < h2 && h1 < h3) {
								flag = false;
							} else { // 1.2G???????????????????????????? 1.2G????????????????????????????
								flag = true;
							}
						} else if (flag2) {
							if (h2 < h1 && h2 < h3) {
								flag = false;
							} else { // 1.2G???????????????????????????? 1.2G????????????????????????????
								flag = true;
							}
						} else {
							if (h3 < h1 && h3 < h2) {
								flag = false;
							} else { // 1.2G???????????????????????????? 1.2G????????????????????????????
								flag = true;
							}
						}
					} else { // ????1.2g????????????????????????????????????????????????????????1.2G????????????????????????
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
	 * ????????????
	 */
	private boolean getCompareresultByTwo(String partmaterial, String partmateria2, String partthickness1,
			String partthickness2, boolean flag1, boolean flag2) {
		boolean flag = false;
		// ????????????????
		if (partthickness1.equals(partthickness2)) { // ????????????????????
			int kn1 = getSheetstrength(partmaterial);
			int kn2 = getSheetstrength(partmateria2);
			// ????????????????1.2G??????????
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
		} else {// ????????????1.2G ??????????????
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
	 * ????????????????
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
	 * ??????????double??????????????0.0
	 */
	private double getDoubleByString(String str) {
		double num = 0.0;
		if (str != null && !str.isEmpty()) {
			num = Double.parseDouble(str);
		}
		return num;
	}

	// ????????????????????excel
	private static void writepicturetosheet(XSSFWorkbook book, XSSFSheet sheet, InputStream is, int colindex,
			int rowindex, int endcolindex, int endrowindex) {
		// ????????????????????????ByteArrayOutputStream????????????ByteArray
		ByteArrayOutputStream byteArrayOut = new ByteArrayOutputStream();
		BufferedImage bufferImg;
		try {
			bufferImg = ImageIO.read(is);
			ImageIO.write(bufferImg, "png", byteArrayOut);
			XSSFDrawing patriarch = sheet.createDrawingPatriarch();
			XSSFClientAnchor anchor = new XSSFClientAnchor(50000, 50000, -50000, -50000, (short) colindex, rowindex,
					(short) endcolindex, endrowindex);
			anchor.setAnchorType(2);
			int picindex = book.addPicture(byteArrayOut.toByteArray(), XSSFWorkbook.PICTURE_TYPE_JPEG);
			// ????????
			patriarch.createPicture(anchor, picindex);
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}

	}

	/*
	 * ????????????3D????
	 */
	private Map<String, File> getAll3DPictures(TCComponentItemRevision blrev) throws TCException {
		Map<String, File> piclist = new HashMap<String, File>();
		TCComponent[] tccdata = blrev.getRelatedComponents("IMAN_3D_snap_shot");
		for (TCComponent tcc : tccdata) {
			String objectname = Util.getProperty(tcc, "object_name");
			// ?????????? ??????????????
			if (Util.isNumber(objectname.substring(0, 1))) {
				File file = downLoadPicture1(tcc, "ThumbnailImage");
				if (file != null) {
					piclist.put(objectname, file);
				}
			}
		}
		return piclist;
	}

	/**
	 * ????????????????????
	 * 
	 * @param picDs1
	 * @return
	 */
	public static File downLoadPicture1(TCComponent comp, String pictype) {
		// TODO Auto-generated method stub

		// System.out.println(">>>downLoadPicture");

		TCComponentDataset dataset = null;
		if (comp instanceof TCComponentDataset) {
			dataset = (TCComponentDataset) comp;
		}
		File file = null;
		if (dataset == null) {
			// System.out.println("dataset==null");
			return null;
		}

		System.out.println("downLoadPicture:" + dataset.toString());
		String type = dataset.getType();
		// "Image","JPEG","Bitmap","TIF","GIF"
		if (!"Vis_Snapshot_2D_View_Data".equals(type) && !"SnapShotViewData".equals(type) && !"Image".equals(type)
				&& !"JPEG".equals(type) && !"Bitmap".equals(type) && !"TIF".equals(type) && !"GIF".equals(type)) {
			// System.out.println("????????????????"+type);
			return null;
		}

		TCComponentTcFile[] files;
		try {

			files = dataset.getTcFiles();
			TCComponent pic = dataset.getNamedRefComponent(pictype);
			String modelname = pic.getProperty("file_name");
			if (files == null || files.length <= 0) {
				return null;
			}
			for (int i = 0; i < files.length; i++) {
				String fileName = files[i].getProperty("file_name");
				System.out.println("fileName:" + fileName);
				if (modelname.equals(fileName)) {
					if (fileName.toLowerCase().endsWith("png") || fileName.toLowerCase().endsWith("jpeg")
							|| fileName.toLowerCase().endsWith("jpg") || fileName.toLowerCase().endsWith("bmp")
							|| fileName.toLowerCase().endsWith("tif") || fileName.toLowerCase().endsWith("gif")) {
						file = files[i].getFmsFile();
						// System.out.println("fms file:"+file.getAbsolutePath());
						return file;
					}
				}

			}
		} catch (TCException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}

		return file;
	}

	// ????????????????????excel
	private static void writepicturetosheet(XSSFWorkbook book, XSSFSheet sheet, BufferedImage bufferImg, int rowindex,
			int colindex, int rowindex2, int colindex2) {
		// ????????????????????????ByteArrayOutputStream????????????ByteArray
		ByteArrayOutputStream byteArrayOut = new ByteArrayOutputStream();
		try {
			ImageIO.write(bufferImg, "png", byteArrayOut);
			XSSFDrawing patriarch = sheet.createDrawingPatriarch();
			XSSFClientAnchor anchor = new XSSFClientAnchor(50000, 50000, -50000, -50000, (short) colindex, rowindex,
					(short) colindex2, rowindex2);
			anchor.setAnchorType(2);
			int picindex = book.addPicture(byteArrayOut.toByteArray(), XSSFWorkbook.PICTURE_TYPE_JPEG);
			// ????????
			patriarch.createPicture(anchor, picindex);
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}

	}

	// ????????????????????????????????????
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

	// ????????????????????????????????????
	private void setCellFormula(XSSFWorkbook book) {
		List shnamelist = new ArrayList();
		int sheetnum = 0;
		List sheetList = new ArrayList();
		sheetnum = book.getNumberOfSheets();
		for (int i = 0; i < sheetnum; i++) {
			String sheetname = book.getSheetName(i);
			if (sheetname.contains("????")) {
				shnamelist.add(sheetname);
			}
			if ((sheetname.contains("PSW") && !sheetname.contains("????")) || sheetname.contains("RSW????")
					|| sheetname.contains("RSW????")) {
				sheetList.add(sheetname);
			}
		}
		if (shnamelist.size() < 1) {
			return;
		}
//		FormulaEvaluator evl = null;
//		evl = new XSSFFormulaEvaluator(book);
		// ????????sheet????????????
		for (int i = 0; i < shnamelist.size(); i++) {
			String shname = (String) shnamelist.get(i);
			XSSFSheet sheet = book.getSheet(shname);
			for (int j = 0; j < 9; j++) {
				XSSFRow row = sheet.getRow(17 + 2 * j);
				if (row == null) {
					row = sheet.createRow(17 + 2 * j);
				}
				XSSFCell cell;
				// CZ??DD??DH??FZ??
				String formula4 = "IF(ISBLANK(DM" + (18 + 2 * j) + "),\"\",DM" + (18 + 2 * j) + ")";// FZ
				cell = row.getCell(181);
				cell.setCellFormula(formula4);
				// evl.evaluateFormulaCell(cell);
				// GA??GB??GC??
				String colname = "DM" + (18 + 2 * j);
				String formula5 = getExcelFormula(sheetList, colname, 6);
				if (!formula5.isEmpty()) {
					cell = row.getCell(182);
					cell.setCellFormula(formula5);
					// evl.evaluateFormulaCell(cell);
				}
				String formula6 = getExcelFormula(sheetList, colname, 32);
				if (!formula6.isEmpty()) {
					cell = row.getCell(183);
					cell.setCellFormula(formula6);
					// evl.evaluateFormulaCell(cell);
				}
				String formula7 = getExcelFormula(sheetList, colname, 58);
				if (!formula7.isEmpty()) {
					cell = row.getCell(184);
					cell.setCellFormula(formula7);
					// evl.evaluateFormulaCell(cell);
				}
				String formula1 = "IF(ISBLANK(DM" + (18 + 2 * j) + "),\"\",GA" + (18 + 2 * j) + ")";// CZ
				cell = row.getCell(103);
				cell.setCellFormula(formula1);
				// evl.evaluateFormulaCell(cell);
				String formula2 = "IF(ISBLANK(DM" + (18 + 2 * j) + "),\"\",GB" + (18 + 2 * j) + ")";// DD
				cell = row.getCell(107);
				cell.setCellFormula(formula2);
				// evl.evaluateFormulaCell(cell);
				String formula3 = "IF(ISBLANK(DM" + (18 + 2 * j) + "),\"\",GC" + (18 + 2 * j) + ")";// DH
				cell = row.getCell(111);
				cell.setCellFormula(formula3);
				// evl.evaluateFormulaCell(cell);

			}
			sheet.setForceFormulaRecalculation(true);
		}
		book.setForceFormulaRecalculation(true);

	}

	/*
	 * ????PSW??RSWsheet????
	 * 
	 * @sheetname sheet????????
	 * 
	 * @colname ????????
	 * 
	 * @colnum ??????????
	 */
	private String getExcelFormula(List sheetname, String colname, int colnum) {
		String formula = "";
		String commonstr = "";
		if (sheetname != null && sheetname.size() > 0) {
			for (int i = 0; i < sheetname.size(); i++) {
				String name = (String) sheetname.get(i);
				if (commonstr.isEmpty()) {
					commonstr = "VLOOKUP(" + colname + ",'" + name + "'!$I$12:$BQ$47," + colnum + ",0)";
				} else {
					commonstr = "IFERROR(" + commonstr + ",VLOOKUP(" + colname + ",'" + name + "'!$I$12:$BQ$47,"
							+ colnum + ",0))";
				}
			}
			if (!commonstr.isEmpty()) {
				formula = "IFERROR(" + commonstr + ",\"\")";
				formula = "IF(" + formula + "=0,\"\"," + formula + ")";
			}
		}
		System.out.println("??????" + formula);
		return formula;
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
				XSSFCellStyle newstyle = book.createCellStyle();
				newstyle = (XSSFCellStyle) style.clone();
				if(bgcolor > -1)
				{
					newstyle.setFillForegroundColor((short)bgcolor);
					newstyle.setFillPattern(CellStyle.SOLID_FOREGROUND);
				}
				if(colorindex > -1)
				{
					// ????????????
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
					// ????????????
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
	 * ??????????????SOP??
	 */
	private boolean getIsSOPAfter() {
		boolean flag = false;
		ArrayList edition = getEditionSizeRule();
		if (edition != null && edition.size() > 0) {
			if (edition.contains(Edition)) {
				return false;
			}
		}
		if (Edition != null) {
			if (Edition.length() == 1) {
				char c = Edition.charAt(0);
				if (c >= 'A' && c <= 'Z') {
					flag = true;
				}
			}
			if (Edition.length() == 2) {
				char c = Edition.charAt(0);
				char cc = Edition.charAt(1);
				if (c >= 'A' && c <= 'Z' && cc >= 'A' && cc <= 'Z') {
					flag = true;
				}
			}
		}
		return flag;
	}

	/*
	 * ??????????????SOP??
	 */
	private boolean getIsTeSOPAfter(String bc) {
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

	// ????????????????????????????
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
}
