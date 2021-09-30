package com.dfl.report.mfcadd;

import java.io.File;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Collections;
import java.util.Comparator;
import java.util.Date;
import java.util.HashMap;
import java.util.Iterator;
import java.util.List;
import java.util.Map;
import java.util.TreeMap;

import com.dfl.report.common.TCComponentUtils;
import com.dfl.report.util.ReportViwePanel;
import com.dfl.report.util.Util;
import com.teamcenter.rac.aif.kernel.AIFComponentContext;
import com.teamcenter.rac.cme.kernel.bvr.FlowUtil;
import com.teamcenter.rac.cme.kernel.mfg.IMfgFlow;
import com.teamcenter.rac.cme.kernel.mfg.IMfgNode;
import com.teamcenter.rac.kernel.TCComponent;
import com.teamcenter.rac.kernel.TCComponentBOMLine;
import com.teamcenter.rac.kernel.TCComponentDataset;
import com.teamcenter.rac.kernel.TCComponentDatasetType;
import com.teamcenter.rac.kernel.TCComponentItem;
import com.teamcenter.rac.kernel.TCComponentItemRevision;
import com.teamcenter.rac.kernel.TCException;
import com.teamcenter.rac.kernel.TCSession;
import com.teamcenter.rac.kernel.tcservices.TcBOMService;
import com.teamcenter.services.rac.core._2008_06.DataManagement.CreateResponse;

public class EDrawingASummaryReportOperation {
	TCComponentBOMLine bopLine = null;
	TCComponent datasetLocation = null;
	String title = "";
	String curdate = "";
	private SimpleDateFormat dateFormat = new SimpleDateFormat("yyyy.M.dd");
	int rows = 0;
	private List<TCComponentBOMLine> lstPertLines;
	//private List<Boolean> lstLR = null;
	private Map<TCComponentBOMLine, Boolean> mapPertLR;
	private List<NissanStation> lstPerts;
	private HashMap<TCComponentBOMLine, NissanStation> hmLinePert;
	private HashMap<String, NissanStation> hmLineNamePert;
	private TCSession session;
	TreeMap<String, NissanStation> tmSeqStation;
	List<NissanStation> lstFirstLine ;
	List<List<NissanStation>> lstProcessRoute;
	int maxRows = 0;
	String docNo = "";
	String version  = "";
	Map<String, String> mapBodyLineName;
	private Map<TCComponentBOMLine, String> mapStationName;
	public EDrawingASummaryReportOperation(TCComponentBOMLine bop, TCComponent folder, String ver) {
		bopLine = bop;
		session = bop.getSession();
		datasetLocation = folder;
		version = ver;
		lstPertLines = new ArrayList<TCComponentBOMLine>();
		lstPerts = new ArrayList<NissanStation>();
		lstFirstLine = new ArrayList<NissanStation>();
		//lstLR = new ArrayList<Boolean>();
		mapPertLR = new HashMap<TCComponentBOMLine, Boolean>();
		hmLinePert = new HashMap<TCComponentBOMLine, NissanStation>();
		hmLineNamePert = new HashMap<String, NissanStation>();
		tmSeqStation = new TreeMap<String, NissanStation>();
		lstProcessRoute = new ArrayList<List<NissanStation>>();
		mapBodyLineName = new HashMap<String, String>();
		mapLineMulStat = new HashMap<TCComponentBOMLine, Boolean>();
		mapLineName = new HashMap<TCComponentBOMLine, String>();
		mapStationName = new HashMap<TCComponentBOMLine, String>();
		getAndoutReport();
	}
	public void getAndoutReport() {
		try {
			// ������ʾ���Ȳ����ִ�в���
			ReportViwePanel viewPanel = new ReportViwePanel("���ɱ���");
			viewPanel.setVisible(true);
			viewPanel.addInfomation("���ڻ�ȡģ��...\n", 20, 100);
			// ��ѯ������ģ��
			String inputStream = TemplateUtil.getTemplateFile("DFL_Template_ManagementOfA");
//			if (inputStream == null) {
//				viewPanel.addInfomation("����û���ҵ�������ͼA���ģ�壬����ϵϵͳ����Ա����TC�����ģ��(����Ϊ��DFL_Template_ManagementOfA)\n", 100,100);
//				return;
//			}
			String inputStreamS2 = TemplateUtil.getTemplateFile("DFL_Template_ManagementOfASheet2");
//			if(inputStreamS2 == null) {
//				viewPanel.addInfomation("����û���ҵ�������ͼA��·��ͼ��ģ�壬����ϵϵͳ����Ա����TC�����ģ��(����Ϊ��DFL_Template_ManagementOfASheet2)\n", 100,100);
//				return;
//			}
			viewPanel.addInfomation("��ʼ�������...\n", 35, 100);
			
			getReportData(this.bopLine);
			
			List<String[]> lstReportDatas = getShee2ReportData();
			int rows = lstReportDatas.size();
			int i = 0, j = 0;
			int pages = rows % 10 == 0 ? rows /10 : rows/10 + 1;
			
			if(pages == 0) {
				Util.callByPass(session, false);
				viewPanel.addInfomation("δ�鵽·��ͼ��Ϣ���޷����������ͼA��...\n", 100, 100);
				return ;
			}
			
			String familycode = bopLine.getItemRevision().getProperty("project_ids");// ����
			String vehicle = Util.getDFLProjectIdVehicle(familycode);
			String veh = "";
			String factory = "";
			String bopName = bopLine.getItemRevision().getProperty("object_name");
			String[] splits = bopName.split("_");
			String fac = "";
			if(splits.length > 3) {
				veh = splits[1];
				factory = splits[2];
				StringBuffer sb = new StringBuffer();
				char[] facs = factory.toCharArray();
				for(i = 0; i < facs.length ; i ++) {
					if(facs[i] >= 'A' && facs[i] <= 'Z') {
						sb.append(facs[i]);
					}else if(facs[i] >= '1' &&facs[i] <= '9') {
						sb.append(facs[i]);
						break;
					}
				}
				fac = sb.toString();
			}
			title = vehicle + "_������ͼA��";
			SimpleDateFormat sim = new SimpleDateFormat("yyyy��MM��");
			curdate = sim.format(new Date());
			viewPanel.addInfomation("��ʼд���ݣ������ĵȴ�...\n", 70, 100);
			POIExcel poi = new POIExcel();
			poi.specifyTemplate(inputStream, 0);
			poi.fillCellValue(5, 3, "    �������̣�" + fac + "������װ����");
			poi.fillCellValue(6, 3, "    ��    �ͣ�" + vehicle);
			poi.fillCellValue(7, 3, "    ��    �Σ�" + version);
			docNo = vehicle + "-" + factory ;
			poi.fillCellValue(8, 3, "    �ļ���ţ�" + docNo+ "-GG-A");
			poi.fillCellValue(10, 3, "    �������ڣ�" + curdate);
			String username = session.getUser().getProperty("user_name");
			poi.fillCellValue(12, 1, username);
			
			String newName = System.getenv("TMP") + File.separator + MFCUtility.fileNameReplace(title, "") + ".xlsx";
			
			poi.outputExcel(newName);
			poi.close();
			
			File ftmp = new File(inputStream);
			ftmp.delete();
			inputStream = newName;
			
			System.out.println("pages := " + pages);
			poi = new POIExcel();
			poi.specifyTemplate(inputStream, 1);
			poi.fillCellValue(32, 13, String.valueOf(pages + 1));
			poi.fillCellValue(32, 12, String.valueOf(1));
			poi.outputExcel(inputStream);
			poi.close();
			String[] sheetNames = new String[pages];
			sheetNames[0] = "02����";
			if(pages > 1) {
				poi = new POIExcel();
				poi.specifyTemplate(inputStream, 2);
				poi.fillCellValue(1, 2, username);
				for(i = 1; i < pages; i ++) {
					sheetNames[i] = StringUtil.leftStrcat((i + 2) + "", 2, "0");
					sheetNames[i] = sheetNames[i] + "����";
				}
				poi.cloneTemplate(2, sheetNames);
				poi.outputExcel(inputStream);
				poi.close();
			}
			int rowIndex = 0;
			for(i = 0 ; i < pages; i ++) {
				poi = new POIExcel();
				poi.specifyTemplate(inputStream, sheetNames[i]);
				poi.fillCellValue(1, 2, username);
				poi.fillCellValue(0, 8, vehicle +"��װ���̹�����Ŀһ����");
				poi.fillCellValue(1, 1, username);
				poi.fillCellValue(32, 13, String.valueOf(pages + 1));
				poi.fillCellValue(32, 12, String.valueOf(i + 2));
				System.out.println("i := " + i + " -- > " + sheetNames[i]);
				for(j = 0; j < 10; j ++) {
					if(i * 10 + j == rows) {
						break;
					}
					System.out.println("write excel");
					rowIndex = j * 2 + 5;
					String[] rowdata = lstReportDatas.get(i * 10 + j);
					poi.fillCellValue(rowIndex, 1, rowdata[0]);
					poi.fillCellValue(rowIndex, 2, rowdata[1]);
					poi.fillCellValue(rowIndex, 4, rowdata[2]);
					poi.fillCellValue(rowIndex, 5, rowdata[3]);
					poi.fillCellValue(rowIndex, 7, rowdata[4]);
					poi.fillCellValue(rowIndex, 8, rowdata[5]);
					poi.fillCellValue(rowIndex, 9, rowdata[6]);
					poi.fillCellValue(rowIndex, 10, "�ο�������ͼB��");
					poi.fillCellValue(rowIndex, 11, "��");
				}
				try{
					if(i > 0) {
						poi.zoomSheet(2 + i, 33, 13, 70, false);
					}
					poi.outputExcel(inputStream);
				}catch(Exception e) {
					e.printStackTrace();
				}
				poi.close();
			}
			
			viewPanel.addInfomation("�������ݼ��������ĵȴ�...\n", 90, 100);
			TCComponentDatasetType dsType = (TCComponentDatasetType) bopLine.getSession().getTypeComponent("MSExcelX");
			TCComponentDataset dataset = dsType.create(title, "", "MSExcelX");
			dataset.setFiles(new String[]{ inputStream }, new String[]{ "excel" });
//			if(datasetLocation instanceof TCComponentFolder) {
//				datasetLocation.add("contents", dataset);
//			}else if(datasetLocation instanceof TCComponentItemRevision) {
//				datasetLocation.add("IMAN_specification", dataset);
//			}
			poi = new POIExcel();
			poi.specifyTemplate(inputStreamS2);
			int paths = this.lstProcessRoute.size();
			System.out.println("paths := " + paths);
			if(paths > 3) {
				for(i = 0; i < paths - 3; i ++) {
					poi.copyCell(0, 0, 0, 6 + i * 2, 6 + i * 2, true);
					poi.copyCell(0, 1, 0, 7 + i * 2, 7 + i * 2, true);
					for(j = 1; j < 9; j ++) {
						poi.copyCell(1, 0, j, 6 + i * 2, 6 + i * 2, true);
						poi.copyCell(1, 1, j, 7 + i * 2, 7 + i * 2, true);
					}
				}
			}
			System.out.println("maxRows := " + maxRows);
			if(this.maxRows > 8) {
				poi.appendRow(8, this.maxRows - 8);
			}
			for(i = 0; i < paths; i ++) {
				List<NissanStation> list = this.lstProcessRoute.get(i);
				rows = list.size();
				System.out.println("rows := " + rows);
				for(j = 0; j <rows; j ++) {
					NissanStation station = list.get(j);
					String sname = station.getName();
					//sname = this.getStationName(station.getCurLine(), sname);
					poi.fillCellValue(j + 1, i * 2 , sname);
					poi.fillCellValue(j + 1, i * 2 + 1 , station.getSeqno() + "");
				}
			}
			String newName2 = System.getenv("TMP") + File.separator + MFCUtility.fileNameReplace(title + "-������ͼ·��ͼ", "") + ".xlsx";
			poi.outputExcel(newName2);
			poi.close();
			File ftmp2 = new File(inputStreamS2);
			ftmp2.delete();
			inputStreamS2 = newName2;
			TCComponentDataset ljtDataset = dsType.create(title + "-������ͼ·��ͼ", "", "MSExcelX");
			ljtDataset.setFiles(new String[]{ inputStreamS2 }, new String[]{ "excel" });
			
			saveFiles(dataset, ljtDataset);
			File file = new File(inputStream);
			file.delete();
			file = new File(inputStreamS2);
			file.delete();
			viewPanel.addInfomation("���������ɣ����ں�װ�������հ汾�����²鿴��...\n", 100, 100);
		}catch(Exception e) {
			e.printStackTrace();
			MFCUtility.errorMassges("�쳣��" + e.getLocalizedMessage());
		}finally {
			try {
				Util.callByPass(session, false);
			} catch (TCException e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			}
		}
	}
	private HashMap<TCComponentBOMLine, Boolean> mapLineMulStat;
	private HashMap<TCComponentBOMLine, String> mapLineName;
	private String getStationName(TCComponentBOMLine statLine, String statName) {
		if(this.mapStationName.containsKey(statLine)) {
			return mapStationName.get(statLine);
		}
		String name = statName;
		try {
			TCComponentBOMLine lineLine = statLine.parent();
			if(mapLineMulStat.containsKey(lineLine)) {
				boolean mulStat = mapLineMulStat.get(lineLine);
				String lineName = mapLineName.containsKey(lineLine) ? mapLineName.get(lineLine) : "";
				if(mulStat) {
					if(StringUtil.isEmpty(lineName)) {
						mapStationName.put(statLine, name);
						return name;
					}
					mapStationName.put(statLine, lineName + " " + statName);
					return lineName + " " + statName;
				}else {
					mapStationName.put(statLine, lineName);
					return lineName;
				}
			}
		}catch(Exception e) {
			e.printStackTrace();
		}
		return name;
	}
	/**
	 * ��ȡ��������
	 * @return
	 */
	private List<String[]> getShee2ReportData(){
		List<String[]> lstReports = new ArrayList<String[]>();
	    Iterator<String> itSeq = this.tmSeqStation.keySet().iterator();
		while(itSeq.hasNext()) {
			String seqno = itSeq.next();
			NissanStation station = this.tmSeqStation.get(seqno);
			try {
				TCComponentBOMLine curLine = station.getCurLine();
				String jobName = station.getName();//curLine.getPropertyDisplayableValue("bl_rev_object_name");
				String jobContent = "";
				String b8_WPImptLevel = "";
				if(curLine.getItem().getType().equals("B8_BIWMEProcStat")) {
					//jobName = getStationName(curLine, jobName);//mifc
					jobContent = curLine.getItemRevision().getProperty("b8_OperationContent");
					b8_WPImptLevel = curLine.getItemRevision().getProperty("b8_WPImptLevel");
					if(station.isLeftRight() && !StringUtil.isEmpty(b8_WPImptLevel)) {
						b8_WPImptLevel = b8_WPImptLevel + "/" + b8_WPImptLevel;
					}
					System.out.println(curLine.getItemRevision() + " --> " + curLine.getItemRevision().getType() + " --> " + b8_WPImptLevel);
				}else {
					b8_WPImptLevel = curLine.getItemRevision().getProperty("b8_TorqueImptLevel");
					System.out.println(curLine.getItemRevision() + " --> " + curLine.getItemRevision().getType() + " --> " + b8_WPImptLevel);
					jobContent = jobName;
				}
				String equipment = MFCUtility.getEquipmentByJobName(jobContent);
				String mgrItems = MFCUtility.getMgrItemsByJobName(jobContent);
				String pinbao = "";
				String zhizao = "";
				if(mgrItems.contains("A") || mgrItems.contains("B") || mgrItems.contains("C") || mgrItems.contains("D") || mgrItems.contains("E") 
						|| mgrItems.contains("F") || mgrItems.contains("G") || mgrItems.contains("H") || mgrItems.contains("I") || mgrItems.contains("O")) {
					pinbao = "��";
					zhizao = "��";
				}
				if(mgrItems.contains("J") || mgrItems.contains("K")|| mgrItems.contains("L")|| mgrItems.contains("M")|| mgrItems.contains("N")) {
					pinbao = "��";
					zhizao = "��";
				}
				if(b8_WPImptLevel == null) {
					b8_WPImptLevel = "";
				}
				lstReports.add(new String[] {station.getSeqno() + "", jobName, equipment, mgrItems, b8_WPImptLevel, zhizao, pinbao});
			}catch(Exception e) {
				e.printStackTrace();
			}
		}
		return lstReports;
	}
	/**
	 * ��ȡ��������
	 * @param pline
	 */
	private void getReportData(TCComponentBOMLine pline) {
		getPertRelations();
		makeTheProcessRoute();
	}
	/**
	 * �ж��Ƿ���ǰ���������Ƿ���·����
	 * @param pertLines
	 * @return
	 */
	boolean hasPreOrSus(TCComponentBOMLine[] pertLines) {
		try {
			MFCUtility.loadProperties(session, pertLines, new String[] {"Mfg0predecessors", "Mfg0successors", "bl_rev_object_name"});
			for(TCComponentBOMLine pertLine : pertLines) {
				TCComponent[] Mfg0predecessors = pertLine.getReferenceListProperty("Mfg0predecessors");
				if(Mfg0predecessors != null && Mfg0predecessors.length > 0) {
					return true;
				}
				TCComponent[] Mfg0successors = pertLine.getReferenceListProperty("Mfg0successors");
				if(Mfg0successors != null && Mfg0successors.length > 0) {
					return true;
				}
				List<IMfgFlow> listSuccessors = FlowUtil.getScopeOutputFlows(pertLine);//�ⲿ������λ
				if(listSuccessors != null && listSuccessors.size() > 0) {
					return true;
				}
				List<IMfgFlow> listPredecessors = FlowUtil.getScopeInputFlows(pertLine);//�ⲿǰ����λ
				if(listPredecessors != null && listPredecessors.size() > 0) {
					return true;
				}
			}
		}catch(Exception e) {
			e.printStackTrace();
		}
		return false;
	}
	//��ȡ�����й�λ�������ǰ���ϵ
	private void getPertRelations() {
		//������·���ϵ�ÿ���㼯�ϵ�һ��
		getPertBOMLines();//��ȡ��ǰ���ϵ��BOMLine
		//�����ϵ�һ��Ľڵ㣬��ȡ��ǰ�����������ɽڵ�����Լ����� ���ڵ�
		TCComponentBOMLine[] pertLines = this.lstPertLines.toArray(new TCComponentBOMLine[0]);
		//Boolean[] leftRight = this.lstLR.toArray(new Boolean[0]);
		int index = 0; 
		System.out.println("pertLines.len := " + pertLines.length);
		MFCUtility.loadProperties(session, pertLines, new String[] {"Mfg0predecessors", "Mfg0successors", "bl_rev_object_name"});
		for(TCComponentBOMLine pertLine : pertLines) {
			NissanStation station = new NissanStation();
			station.setCurLine(pertLine);
			Boolean lr =this.mapPertLR.containsKey(pertLine) ? this.mapPertLR.get(pertLine) : false;
			try{
				station.setLeftRight(lr);
			}catch(Exception e) {
				e.printStackTrace();
			}
			System.out.println("pertLine := " + pertLine);
			String stationName = "";
			try {
				List<IMfgFlow> listSuccessors = FlowUtil.getScopeOutputFlows(pertLine);//�ⲿ������λ
				List<IMfgFlow> listPredecessors = FlowUtil.getScopeInputFlows(pertLine);//�ⲿǰ����λ
				TCComponent[] Mfg0predecessors = pertLine.getReferenceListProperty("Mfg0predecessors");//�ڲ�ǰ��
				TCComponent[] Mfg0successors = pertLine.getReferenceListProperty("Mfg0successors");//�ڲ�����
				stationName = pertLine.getPropertyDisplayableValue("bl_rev_object_name");
				if(pertLine.getItem().getType().equals("B8_BIWMEProcStat")) {
//					String pname = pertLine.parent().getPropertyDisplayableValue("bl_B8_BIWMEProcStatRevision_b8_ChineseName");
//					if(pname == null) {
//						pname = "";
//					}
//					station.setName(pname + " " + stationName);
					station.setName(this.getStationName(pertLine, stationName));
				}else {
					station.setName(stationName);
				}
				System.out.println("Ŀ�깤λ��" + pertLine + " LeftRight = " + lr);
				if(lr) {
					station.setName(station.getName().replace("��", "").replace("��", "") + " ��/��");
				}
				if(Mfg0successors == null || Mfg0successors.length == 0) {
					if(listSuccessors != null && listSuccessors.size() > 0) {
						TCComponentBOMLine[] successors = new TCComponentBOMLine[listSuccessors.size()];
						int idx = 0;
						for (IMfgFlow flow : listSuccessors) {
							IMfgNode node = flow.getSuccessor();
							TCComponentBOMLine sucComp = (TCComponentBOMLine) node.getComponent();
							successors[idx] = sucComp;
							idx ++;
						}
						station.setMfg0successors(successors);
					}else {
						station.setMfg0successors(null);
					}
					
				}else {
					TCComponentBOMLine[] successors = new TCComponentBOMLine[Mfg0successors.length];
					int idx = 0;
					for(TCComponent comp : Mfg0successors) {
						if(comp instanceof TCComponentBOMLine) {
							successors[idx] = (TCComponentBOMLine)comp;
							idx ++;
						}
					}
					station.setMfg0successors(successors);
				}
				if(Mfg0predecessors == null || Mfg0predecessors.length == 0) {
					if(listPredecessors != null && listPredecessors.size() > 0) {
						TCComponentBOMLine[] predecessors = new TCComponentBOMLine[listPredecessors.size()];
						int idx = 0;
						for (IMfgFlow flow : listPredecessors) {
							IMfgNode node = flow.getPredecessor();
							TCComponentBOMLine sucComp = (TCComponentBOMLine) node.getComponent();
							predecessors[idx] = sucComp;
							idx ++;
						}
						station.setMfg0predecessors(predecessors);
					}else {
						station.setMfg0predecessors(null);
						//lstFirstLine.add(station);
					}
				}else {
					TCComponentBOMLine[] predecessors = new TCComponentBOMLine[Mfg0predecessors.length];
					List<TCComponentBOMLine> lists = new ArrayList<TCComponentBOMLine>();
					for(TCComponent comp : Mfg0predecessors) {
						if(comp instanceof TCComponentBOMLine) {
							lists.add((TCComponentBOMLine)comp);
						}
					}
					if(listPredecessors != null && listPredecessors.size() > 0) {
						for (IMfgFlow flow : listPredecessors) {
							IMfgNode node = flow.getPredecessor();
							TCComponentBOMLine sucComp = (TCComponentBOMLine) node.getComponent();
							lists.add(sucComp);
						}
					}
					predecessors = lists.toArray(new TCComponentBOMLine[0]);
					station.setMfg0predecessors(predecessors);
				}
				if(station.getMfg0predecessors() == null && station.getMfg0successors() == null) {
					System.out.println(pertLine + " --> ��ȫû��ǰ��������������·����.......");
					continue;
				}
				if(station.getMfg0predecessors() == null){
					lstFirstLine.add(station);
				}
				this.lstPerts.add(station);
				this.hmLinePert.put(pertLine, station);
				this.hmLineNamePert.put(stationName, station);
			} catch (Exception e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			}
			index ++;
		}
		//
		System.out.println("lstFirstLine.size ;= " + lstFirstLine.size());
		
	}
	private int curSeq = 0;
	private void makeTheProcessRoute() {
		if(lstFirstLine == null || lstFirstLine.size() == 0	) {
			System.out.println("û�пյģ�lstFirstLine == null || lstFirstLine.size() == 0	");
			return;
		}
		int i = 0;
		int firsts = this.lstFirstLine.size();
		curSeq = 0;
		//�������п�ͷ�ڵ��������ϱ��������ɱ��
		for( i = 0; i < firsts; i ++) {
			NissanStation first = this.lstFirstLine.get(i);
//			first.setSeqno(1);
//			this.tmSeqStation.put(StringUtil.leftStrcat("1", 4, "0"), first);
			if(first.getSeqno() == 0) {
				generateSeqno4Station(first);//���ɱ��
			}
		}
		/**
		 * �����п�ͷ�����������
		 */
		Comparator comparator = new Comparator() {
			public int compare(Object paramObject1, Object paramObject2) {
				NissanStation station1 = (NissanStation) paramObject1;
				NissanStation station2 = (NissanStation) paramObject2;
				int seq1 = station1.getSeqno();
				int seq2 = station2.getSeqno();
				return seq1 - seq2;
			}
		};
		Collections.sort(lstFirstLine,comparator);
		Iterator<String> itSeq = this.tmSeqStation.keySet().iterator();
		while(itSeq.hasNext()) {
			String seqno = itSeq.next();
			System.out.println("seqno := " + seqno + " --> " + this.tmSeqStation.get(seqno).getCurLine());
		}
		/**
		 * ����ÿ��·������·��ͼ
		 */
		for(i = 0; i < firsts; i ++) {
			NissanStation station = this.lstFirstLine.get(i);
			System.out.println("station.seqno := " + station.getSeqno());
			List<NissanStation> list = new ArrayList<NissanStation>();
			station.setPassed(true);
			list.add(station);
			generationPath4Report(station, list);//��ȡ����·��
		}
	}
	/**
	 * ������һ·���£�����˳������ͼ
	 * @param station
	 * @param listBef
	 */
	private void generationPath4Report(NissanStation station, List<NissanStation> listBef) {
		TCComponentBOMLine[] successors = station.getMfg0successors();
		if(successors != null && successors.length > 0	) {
			NissanStation nextStation = this.hmLinePert.get(successors[0]);
			if(nextStation == null) {
				System.out.println(successors[0] + " --> no station , 3333333333");
				try {
					String name = successors[0].getPropertyDisplayableValue("bl_rev_object_name");
					nextStation = this.hmLineNamePert.get(name);
				} catch (Exception e) {
					// TODO Auto-generated catch block
					e.printStackTrace();
				}
				if(nextStation == null) {
					this.lstProcessRoute.add(listBef);
					if(listBef.size() > maxRows) {
						maxRows = listBef.size();
					}
					return;
				}
			}
			if(nextStation.isPassed()) {
				listBef.add(nextStation);
				this.lstProcessRoute.add(listBef);
				if(listBef.size() > maxRows) {
					maxRows = listBef.size();
				}
			}else {
				nextStation.setPassed(true);
				listBef.add(nextStation);
				generationPath4Report(nextStation, listBef);
			}
//			for(i = 0; i < count; i ++) {
//				NissanStation nextStation = this.hmLinePert.get(successors[i]);
//				if(i > 0) {
//					List<NissanStation> newList = new ArrayList<NissanStation>();
//					for(int j = 0; j < tmpList.size(); j ++) {
//						newList.add(tmpList.get(j));
//					}
//					newList.add(nextStation);
//					generationPath4Report(nextStation, newList);
//				}else {
//					listBef.add(nextStation);
//					generationPath4Report(nextStation, listBef);
//				}
//			}
		}else {
			this.lstProcessRoute.add(listBef);
			if(listBef.size() > maxRows) {
				maxRows = listBef.size();
			}
		}
	}
	/**
	 * ����ǰ�����źţ���������ǰ��ģ����Žṹ�����±�������������������
	 * @param station
	 */
	private void generateSeqno4Station(NissanStation station) {
		makePredecessorsPath(station);
		TCComponentBOMLine[] successors = station.getMfg0successors();
		if(successors != null && successors.length > 0	) {
			int i = 0; 
			int count = successors.length;
			for(i = 0; i < count; i ++) {
				NissanStation nextStation = this.hmLinePert.get(successors[i]);
				if(nextStation == null) {
					try {
						String name = successors[i].getPropertyDisplayableValue("bl_rev_object_name");
						nextStation = this.hmLineNamePert.get(name);
					} catch (Exception e) {
						// TODO Auto-generated catch block
						e.printStackTrace();
					}
					if(nextStation == null) {
						continue;
					}
					
				}
				generateSeqno4Station(nextStation);
			}
		}
	}
	/**
	 * ǰ���������
	 * @param station
	 */
	private void makePredecessorsPath(NissanStation station) {
		TCComponentBOMLine[] predecessors = station.getMfg0predecessors();
		if(predecessors != null && predecessors.length > 0) {
			int i = 0; 
			int count = predecessors.length;
			System.out.println("station := " + station.getCurLine() + " --> ǰ��" + count);
			//if(count > 1) {
				for(i = 0; i < count; i ++) {
					NissanStation befStation = this.hmLinePert.get(predecessors[i]);
					if(befStation == null) {
						try {
							String name = predecessors[i].getPropertyDisplayableValue("bl_rev_object_name");
							befStation = this.hmLineNamePert.get(name);
						} catch (Exception e) {
							// TODO Auto-generated catch block
							e.printStackTrace();
						}
						if(befStation == null) {
							continue;
						}
					}
					if(befStation.getSeqno() == 0) {
						makePredecessorsPath(befStation);
					}
				}
			//}
			curSeq ++;
			station.setSeqno(curSeq);
			System.out.println("station := " + station.getCurLine() + " --> " + curSeq);
			this.tmSeqStation.put(StringUtil.leftStrcat(curSeq + "", 4, "0"), station);
		}else if(this.lstFirstLine.contains(station)) {
			curSeq ++;
			station.setSeqno(curSeq);
			this.tmSeqStation.put(StringUtil.leftStrcat(curSeq + "", 4, "0"), station);
			System.out.println("station := " + station.getCurLine() + " --> " + curSeq);
		}
	}
	//��ȡ��pertͼ��bomline
	/**
	 * ����BOP��������Pertͼ���ϵ�Ĺ�λ���ǵ����ߣ����򣨵����ߣ����뵽������
	 */
	private void getPertBOMLines() {
		try {
			TcBOMService.expandOneLevel(session, new TCComponentBOMLine[] {this.bopLine});
			AIFComponentContext[] children = bopLine.getChildren();
			int i = 0;
			int count = children.length;
			List<TCComponentBOMLine> lstProcLines = new ArrayList<TCComponentBOMLine>();
			List<TCComponentBOMLine> lstVirProcLines = new ArrayList<TCComponentBOMLine>();
			TCComponentBOMLine cline = null;
			//��ȡ������ߵ�������
			for(i = 0; i < count; i ++) {
				cline = (TCComponentBOMLine)children[i].getComponent();//�������
				if(cline.getItem().getType().equals("B8_BIWMEProcLine")) {
					lstVirProcLines.add(cline);
				}
			}
			TcBOMService.expandOneLevel(session, lstVirProcLines.toArray(new TCComponentBOMLine[0]));
			//��ȡ��λ�������һ������ʵ�����������µĹ�λ��
			count = lstVirProcLines.size();
			int j = 0;
			int cntReal = 0;
			int k = 0;
			int metals;
			//����������ߵ���һ����
			for(i = 0; i < count; i ++) {
				cline = lstVirProcLines.get(i);
				children = cline.getChildren();
				cntReal = children.length;
				System.out.println("cline.getItem().getType() := " + cline.getItem().getType());
				System.out.println("cline.getItem() name  := " + cline.getPropertyDisplayableValue("bl_rev_object_name"));
				if(cline.getItem().getType().equals("B8_BIWMEProcLine") && cline.getPropertyDisplayableValue("bl_rev_object_name").contains("METAL")) {
					System.out.println("��ȡ���е�����......");
//					AIFComponentContext[] metalStations = cline.getChildren();
//					metals = metalStations.length;
					for(j = 0; j < cntReal; j ++) {
						TCComponentBOMLine metalStationLine = (TCComponentBOMLine)children[j].getComponent();
						System.out.println("metalStationLine type := " + metalStationLine.getItem().getType());
						if(metalStationLine.getItem().getType().equals("B8_BIWMEProcLine")) {
							AIFComponentContext[] stations = metalStationLine.getChildren();
							for(k = 0; k < stations.length; k ++) {
								TCComponentBOMLine staLine = (TCComponentBOMLine)stations[k].getComponent();
								if(staLine.getItem().getType().equals("B8_BIWMEProcStat")) {
									lstProcLines.add(staLine);//Ӧ�ý�����1���������ߣ�����������Ҫ��ȡ��λ����Ĺ����������ȡ��������Ĺ�λ����
									System.out.println("��ȡ�������ߵĹ�λ�����뼯��......");
								}
							}
						}
					}
				}else {
					List<String> lstLineName = new ArrayList<String>();
					List<TCComponentBOMLine>  lstRealLines = new ArrayList<TCComponentBOMLine>();
					for(j = 0; j < cntReal; j ++) {
						TCComponentBOMLine realLineLine = (TCComponentBOMLine)children[j].getComponent();
						if(realLineLine.getItem().getType().equals("B8_BIWMEProcLine")) {
							if(!lstRealLines.contains(realLineLine)) {
								lstRealLines.add(realLineLine);
							}
						}
					}
					MFCUtility.loadProperties(session, lstRealLines.toArray(new TCComponentBOMLine[0]), new String[] {"bl_rev_object_name", "bl_B8_BIWMEProcLineRevision_b8_ChineseName"});
					cntReal = lstRealLines.size();
					for(j = 0; j < cntReal; j ++) {
						TCComponentBOMLine realLineLine = lstRealLines.get(j);
						String lineName = realLineLine.getPropertyDisplayableValue("bl_rev_object_name");
						String lineLR = "";
						if(lineName.endsWith("LH") ) {
							lineLR = lineName.substring(0, lineName.length() - 2);
						}else if(lineName.endsWith("RH") ) {
							lineLR = lineName.substring(0, lineName.length() - 2);
						}else {//�����ң���ֱ��д�뼯��
							lstLineName.add(lineName);
							lstProcLines.add(realLineLine);
							continue;
						}
						if(lstLineName.contains(lineLR)) {
							continue;
						}
						for(k = j + 1; k < cntReal; k ++) {
							TCComponentBOMLine realLine = lstRealLines.get(k);
							String name = realLine.getPropertyDisplayableValue("bl_rev_object_name");
							if(!name.equals(lineName) && name.startsWith(lineLR)) {//�����ҵ����ҵ���һ��
								TCComponentBOMLine toAddLine = this.getLeftRightLine(realLine, realLineLine);
								if(toAddLine != null) {
									lstLineName.add(lineLR);
									lstProcLines.add(toAddLine);
								}
								break;
							}
						}
					}
				}
			}
			TCComponentBOMLine[] parentLines = lstProcLines.toArray(new TCComponentBOMLine[0]);
			TcBOMService.expandOneLevel(session, parentLines);
			count = parentLines.length;
			for(i = 0; i < count; i ++) {
				children = parentLines[i].getChildren();
				boolean lr = false;
				String name = parentLines[i].getPropertyDisplayableValue("bl_rev_object_name");
				if(name.endsWith("LH") || name.endsWith("RH")) {
					lr = true;
				}
				cntReal = children.length;
				int cntStats = 0;
				for(j = 0; j < cntReal; j ++) {
					TCComponentBOMLine line = (TCComponentBOMLine)children[j].getComponent();
					String tarType = line.getItem().getType();
					if(tarType.equals("B8_BIWMEProcStat")) {
						cntStats ++;
						this.lstPertLines.add(line);
						//this.lstLR.add(lr);
						mapPertLR.put(line, lr);
						System.out.println("��λ��" + line + " --> �������ƣ�" + name + " LR := " + lr);
					}else if(tarType.equals("B8_BIWOperation") ) {
						System.out.println(line + " --> is a ��װ����");
						this.lstPertLines.add(line);
						//this.lstLR.add(false);
						mapPertLR.put(line, lr);
					}
				}
				String cname  = parentLines[i].getPropertyDisplayableValue("bl_B8_BIWMEProcLineRevision_b8_ChineseName");
				if(cname != null && cname.length() > 0) {
					this.mapLineName.put(parentLines[i], cname);
				}
				if(cntStats > 1) {
					this.mapLineMulStat.put(parentLines[i], true);
				}else {
					this.mapLineMulStat.put(parentLines[i], false);
				}
			}
		} catch (Exception e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
	}
	/**
	 * ��������ҹ�λ���ж���ѡȡ������
	 * @param left
	 * @param right
	 * @return
	 */
	public TCComponentBOMLine getLeftRightLine(TCComponentBOMLine left, TCComponentBOMLine right) {
		TCComponentBOMLine line = left;
		try {
			AIFComponentContext[] leftchild = left.getChildren();
			List<TCComponentBOMLine> lstLeft = new ArrayList<TCComponentBOMLine>();
			int i = 0;
			int count = leftchild.length;
			for(i = 0; i < count; i ++) {
				TCComponentBOMLine	 cline = (TCComponentBOMLine)leftchild[i].getComponent();
				if(cline.getItem().getType().equals("B8_BIWMEProcStat")) {
					lstLeft.add(cline);
				}
			}
			AIFComponentContext[] rightchild = right.getChildren();
			List<TCComponentBOMLine> lstRight = new ArrayList<TCComponentBOMLine>();
			count = rightchild.length;
			for(i = 0; i < count; i ++) {
				TCComponentBOMLine	 cline = (TCComponentBOMLine)rightchild[i].getComponent();
				if(cline.getItem().getType().equals("B8_BIWMEProcStat")) {
					lstRight.add(cline);
				}
			}
			boolean leftroute = this.hasPreOrSus(lstLeft.toArray(new TCComponentBOMLine[0]));
			boolean rightroute = this.hasPreOrSus(lstRight.toArray(new TCComponentBOMLine[0]));
			System.out.println("���ң�" + left + " --> leftroute := " + leftroute);
			System.out.println("���ң�" + right + " --> rightroute := " + rightroute);
			if(leftroute && rightroute) {
				if(lstRight.size() > lstLeft.size()) {
					return right;
				}else {
					return left;
				}
			}else if(rightroute &&!leftroute) {
				return right;
			}else if(!rightroute && leftroute) {
				return left;
			}else {
				return null;
			}
		}catch(Exception e) {
			e.printStackTrace();
		}
		return line;
	}
	// ���ɵı���
		public void saveFiles(TCComponentDataset ds, TCComponentDataset ds2) {
			try {
				TCComponentItemRevision toprev = this.bopLine.getItemRevision();
				TCComponentItemRevision docRev = null;
				TCComponent[] refs = toprev.getRelatedComponents("IMAN_reference");
				int i = 0;
				int count = refs.length;
				for(i = 0; i < count; i ++) {
					if(refs[i].getType().equals("DFL9MEDocument")) {
						if(refs[i].getProperty("object_name").equals(title)) {
							docRev = ((TCComponentItem)refs[i]).getLatestItemRevision();
							break;
						}
					}
				}
				TCComponent[] specs = toprev.getRelatedComponents("IMAN_specification");
				count = specs.length;
				TCComponentDataset dsLJT = null;
				for(i = 0; i < count; i ++) {
					if(specs[i].getType().equals("MSExcelX") && specs[i].getProperty("object_name").equals("������ͼ·��ͼ")) {
						dsLJT = (TCComponentDataset)specs[i];
						break;
					}
				}
				if(dsLJT != null) {
					toprev.cutOperation("IMAN_specification", new TCComponent[] { dsLJT });
					try {
						dsLJT.delete();
					} catch (Exception e2) {

					}
				}
				toprev.add("IMAN_specification", ds2);
				if (docRev != null) {
					// ������ĵ��µ����ݼ�
					// �Ƴ���ʱ����Ҫ�����з��������Ķ����ҳ��������Ƴ�
					TCComponent[] children = TCComponentUtils.getCompsByRelation(docRev, "IMAN_specification");
					for (TCComponent child : children) {
						if (child.getType().equals("MSExcelX") && child.getProperty("object_name").equals(title)) {
							TCComponentDataset dataset = (TCComponentDataset) child;
							docRev.cutOperation("IMAN_specification", new TCComponent[] { dataset });
							try {
								dataset.delete();
							} catch (Exception e2) {

							}
						}
					}
					// ����ĵ������ݼ��Ĺ�ϵ
					docRev.add("IMAN_specification", ds);

				} else {
					Map<String, Object> itemMap = new HashMap<String, Object>();
					Map<String, Object> itemRevisionMap = new HashMap<String, Object>();
					Map<String, Object> itemRevMasterFormMap = new HashMap<String, Object>();
					itemMap.put("item_id", ""); //$NON-NLS-1$ //$NON-NLS-2$
					itemMap.put("object_name", title); //$NON-NLS-1$
					itemMap.put("object_desc", ""); //$NON-NLS-1$
					itemMap.put("object_type", "DFL9MEDocument"); //$NON-NLS-1$
					//itemMap.put("dfl9_vehiclePlant", docNo); 
					itemRevisionMap.put("object_type", "DFL9MEDocumentRevision"); //$NON-NLS-1$
					itemRevisionMap.put("object_name", title); //$NON-NLS-1$
					itemRevisionMap.put("dfl9_process_type", "H"); //$NON-NLS-1$
					itemRevisionMap.put("dfl9_process_file_type", "GG-A"); //$NON-NLS-1$
					itemRevMasterFormMap.put("object_type", "DFL9MEDocumentRevisionMaster"); //$NON-NLS-1$
					CreateResponse respose = TCComponentUtils.create(itemMap, itemRevisionMap, itemRevMasterFormMap);
					int num = respose.serviceData.sizeOfCreatedObjects();
					System.out.println("num := " + num);
					TCComponentItemRevision rev = null;
					TCComponentItem tccomponentitem = null;
					if(num > 0){
						for(i=0;i<num;i++){
							TCComponent comp = respose.serviceData.getCreatedObject(i);
							if(comp instanceof TCComponentItemRevision){
								rev = (TCComponentItemRevision) comp;	
								tccomponentitem = rev.getItem();
							}else if(comp instanceof TCComponentItem) {
								tccomponentitem = (TCComponentItem)comp;
								rev  = tccomponentitem.getLatestItemRevision();
							}
						}
					}
//					TCComponentItemType tcccomponentitemtype = (TCComponentItemType) session
//							.getTypeComponent("DFL9MEDocument");
//					TCComponentItem tccomponentitem = tcccomponentitemtype.create("", "", "DFL9MEDocument", title,
//							"desc", null);
//					tccomponentitem.setProperty("dfl9_process_type", "H");
//					tccomponentitem.setProperty("dfl9_process_file_type", "GG-A");
//					TCComponentItemRevision rev = tccomponentitem.getLatestItemRevision();
//					rev.setProperty("dfl9_vehiclePlant", docNo);
					// ����ĵ������ݼ��Ĺ�ϵ
					rev.add("IMAN_specification", ds);
					// ��Ӻ�װ��λ���ĵ��Ĺ�ϵ
					toprev.add("IMAN_reference", tccomponentitem);
					tccomponentitem.setProperty("dfl9_vehiclePlant", docNo);
				}
			} catch (Exception e) {
				e.printStackTrace();
			}
		}
}
