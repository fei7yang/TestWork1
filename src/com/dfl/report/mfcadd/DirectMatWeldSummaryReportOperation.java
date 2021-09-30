package com.dfl.report.mfcadd;

import java.io.File;
import java.math.BigDecimal;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import com.dfl.report.common.TCComponentUtils;
import com.dfl.report.home.OpenHomeDialog;
import com.dfl.report.util.ReportViwePanel;
import com.dfl.report.util.Util;
import com.teamcenter.rac.aif.kernel.AIFComponentContext;
import com.teamcenter.rac.aifrcp.AIFUtility;
import com.teamcenter.rac.kernel.TCComponent;
import com.teamcenter.rac.kernel.TCComponentBOMLine;
import com.teamcenter.rac.kernel.TCComponentDataset;
import com.teamcenter.rac.kernel.TCComponentDatasetType;
import com.teamcenter.rac.kernel.TCComponentFolder;
import com.teamcenter.rac.kernel.TCComponentItem;
import com.teamcenter.rac.kernel.TCComponentItemRevision;
import com.teamcenter.rac.kernel.TCComponentMEOP;
import com.teamcenter.rac.kernel.TCSession;
import com.teamcenter.rac.kernel.tcservices.TcBOMService;
import com.teamcenter.services.rac.core._2008_06.DataManagement.CreateResponse;

public class DirectMatWeldSummaryReportOperation {
	TCComponentBOMLine bopLine = null;
	TCComponent datasetLocation = null;
	String title = "";
	String curdate = "";
	private SimpleDateFormat dateFormat = new SimpleDateFormat("yyyy.M.dd");
	int rows = 0;
	private List<String> lstBodies = null;
	private HashMap<String, List<String[]>> hmBodyData;
	TCComponentBOMLine[] virtualLines = null;
	TCSession session = null;
	private HashMap<String, String> hmTypeProportion;
	private HashMap<String, String> hmTypePrice;
	private HashMap<String, String> hmTypeNes;
	private HashMap<String, String> hmNesType;
	private HashMap<String, String> hmNesProportion;
	private HashMap<String, String> hmNesPrice;
	private final int COL_TYPE = 0;
	private final int COL_NESTYPE = 1;
	private final int COL_PROPORTION = 2;
	private final int COL_PRICE = 4;
	private HashMap<TCComponentBOMLine, String> hmLineBody;
	String version = "";
	Map<String, String> mapBodyLineName;
	private static final String[] weldProps = new String[] {"b8_modelno", "b8_Long", "b8_LongUOM", "b8_Diameter", "b8_Hight"};
	public DirectMatWeldSummaryReportOperation(TCComponentBOMLine bop, TCComponentBOMLine[] lines, 
			TCComponent folder, String ver) {
		bopLine = bop;
		session = bopLine.getSession();
		virtualLines = lines;
		datasetLocation = folder;
		version = ver;
		lstBodies = new ArrayList<String>();
		hmTypeProportion = new HashMap<String, String>();
		hmTypePrice = new HashMap<String, String>();
		hmTypeNes = new HashMap<String, String>();
		hmNesProportion = new HashMap<String, String>();
		hmNesPrice = new HashMap<String, String>();
		hmNesType = new HashMap<String, String>();
		hmBodyData = new HashMap<String,List< String[]>>();
		hmLineBody = new HashMap<TCComponentBOMLine, String>();
		mapBodyLineName = new HashMap<String, String>();
		getAndoutReport();
	}
	public void getAndoutReport() {
		try {
			// 界面显示进度并输出执行步骤
			ReportViwePanel viewPanel = new ReportViwePanel("生成报表");
			viewPanel.setVisible(true);
			viewPanel.addInfomation("正在获取模板...\n", 20, 100);
			String prefValue = session.getPreferenceService().getStringValue("DFL9_DirectMate_CountRule");
//			if(prefValue == null || prefValue.length() == 0) {
//				viewPanel.addInfomation("错误：配置直材计算规则Excel文件的首选项DFL9_DirectMate_CountRule 未配置\n", 100,100);
//				return;
//			}
			String countRule = TemplateUtil.getTemplateFile(prefValue);
//			if (countRule == null) {
//				viewPanel.addInfomation("错误：没有找到直材计算规则的Excel文件，请先在TC中添加，名称为：" + prefValue, 100,100);
//				return;
//			}
			LargeExcelFileReadUtil example = new LargeExcelFileReadUtil();
			String[][] infos = example.getExcelDatas(countRule);
			try {
				if(infos != null && infos.length > 1) {
					for(int i = 1; i < infos.length ; i ++) {
						if(infos[i].length >0 &&infos[i][this.COL_TYPE] != null && infos[i][this.COL_TYPE].length() > 0) {
							if(infos[i].length > 4 && infos[i][this.COL_PRICE] != null && infos[i][this.COL_PRICE].length() > 0) {
								this.hmTypePrice.put(infos[i][this.COL_TYPE], infos[i][this.COL_PRICE]);
							}
							if(infos[i].length > 2 && infos[i][this.COL_PROPORTION] != null && infos[i][this.COL_PROPORTION].length() > 0) {
								this.hmTypeProportion.put(infos[i][this.COL_TYPE], infos[i][this.COL_PROPORTION]);
							}
							if(infos[i].length > 1 && infos[i][this.COL_NESTYPE] != null && infos[i][this.COL_NESTYPE].length() > 0) {
								this.hmTypeNes.put(infos[i][this.COL_TYPE], infos[i][this.COL_NESTYPE]);
							}
						}
						
						if(infos[i].length > 1 && infos[i][this.COL_NESTYPE] != null && infos[i][this.COL_NESTYPE].length() > 0) {
							if(infos[i].length > 4 &&infos[i][this.COL_PRICE] != null && infos[i][this.COL_PRICE].length() > 0) {
								this.hmNesPrice.put(infos[i][this.COL_NESTYPE], infos[i][this.COL_PRICE]);
							}
							if(infos[i].length > 2 && infos[i][this.COL_PROPORTION] != null && infos[i][this.COL_PROPORTION].length() > 0) {
								this.hmNesProportion.put(infos[i][this.COL_NESTYPE], infos[i][this.COL_PROPORTION]);
							}
							if(infos[i].length >0 &&infos[i][this.COL_TYPE] != null && infos[i][this.COL_TYPE].length() > 0) {
								this.hmNesType.put(infos[i][this.COL_NESTYPE], infos[i][this.COL_TYPE]);
							}
						}
					}
				}
			}catch(Exception e) {
				e.printStackTrace();
			}
			
			
			// 查询并导出模板
			String inputStream = TemplateUtil.getTemplateFile("DFL_Template_HZDirectMetaList");
//			if (inputStream == null) {
//				viewPanel.addInfomation("错误：没有找到直材清单（焊技）的模板，请先在TC中添加模板(名称为：DFL_Template_HZDirectMetaList)\n", 100,100);
//				return;
//			}
			viewPanel.addInfomation("开始输出报表...\n", 35, 100);
			String familycode = bopLine.getItemRevision().getProperty("project_ids");// 车型
			String vehicle = Util.getDFLProjectIdVehicle(familycode);
			String factory = "";
			String bopName = bopLine.getItemRevision().getProperty("object_name");
			String[] splits = bopName.split("_");
			String fac = "";
			String line = "";
			if(splits.length > 3) {
				factory = splits[2];
				char[] facs = factory.toCharArray();
				StringBuffer sb = new StringBuffer();
				for(int i = 0; i < facs.length; i ++	) {
					if(facs[i] >= 'A' && facs[i] <= 'Z') {
						sb.append(facs[i]);
					}else if(facs[i]>='1' && facs[i] <= '9') {
						switch(facs[i]) {
						case '1':
							sb.append("一") ;
							break;
						case '2':
							sb.append("二") ;
							break;
						case '3':
							sb.append("三") ;
							break;
						case '4':
							sb.append("四") ;
							break;
						case '5':
							sb.append("五") ;
							break;
						case '6':
							sb.append("六") ;
							break;
						case '7':
							sb.append("七") ;
							break;
						case '8':
							sb.append("八") ;
							break;
						case '9':
							sb.append("九") ;
							break;
						}
						break;
					}
				}
				fac = sb.toString();
				line = splits[2].substring(fac.length());
				if(line.contains("1")) {
					line = "一";
				}else if(line.contains("2")) {
					line = "二";
				}else if(line.contains("3")) {
					line = "三";
				}else if(line.contains("4")) {
					line = "四";
				}else if(line.contains("5")) {
					line = "五";
				}else if(line.contains("6")) {
					line = "六";
				}else if(line.contains("7")) {
					line = "七";
				}else if(line.contains("8")) {
					line = "八";
				}else if(line.contains("9")) {
					line = "九";
				}
				
			}
			
			
			title = vehicle + "_" + version + "_直材清单（焊技）_";
			SimpleDateFormat sim = new SimpleDateFormat("yyyy.MM.dd HH时");
			SimpleDateFormat sim2 = new SimpleDateFormat("yyyy年MM月");
			SimpleDateFormat sim3 = new SimpleDateFormat("yyyy.M");
			curdate = sim.format(new Date());
			title = title + curdate;
			getReportData(this.bopLine);
			if(this.hmBodyData.size() == 0) {
				viewPanel.addInfomation("未在所选目标下找到直材信息...\n", 100, 100);
				try {
					viewPanel.setVisible(false);
					viewPanel.dispose();
				}catch(Exception e) {
					e.printStackTrace();
				}
				MFCUtility.errorMassges("未在所选目标下找到直材信息 ！");
				return;
			}
			String newName = System.getenv("TMP") + File.separator + MFCUtility.fileNameReplace(title, "") + ".xlsx";
			
			viewPanel.addInfomation("开始写数据，请耐心等待...\n", 70, 100);
			POIExcel poi = new POIExcel();
			poi.specifyTemplate(inputStream, 0);
			poi.fillCellValue(5, 3, "      工厂工程：" + fac + "工厂焊装工程");
			poi.fillCellValue(6, 3, "      车    型：共通");
			poi.fillCellValue(7, 3, "      版    次：" + version);
			poi.fillCellValue(8, 3, "      文件编号：共通-" + factory + "-ZC");
			poi.fillCellValue(10, 3, "      编制日期：" + sim2.format(new Date()));
			String username = session.getUser().getProperty("user_name");
			poi.fillCellValue(12, 1, username);
			poi.outputExcel(newName);
			poi.close();
			File ftmp = new File(inputStream);
			ftmp.delete();
			inputStream = newName;
			int i = 0;
			int j = 0, k = 0;
			int count = this.lstBodies.size();
			List<String[]> list = new ArrayList<String[]>();
			for(i = 0 ; i < count; i ++) {
				List<String[]> orderlist = this.hmBodyData.get(lstBodies.get(i));
				if(list == null) {
					System.out.println("error body data : = " + lstBodies.get(i));
					continue;
				}
				System.out.println(lstBodies.get(i) + " --> " + orderlist.size());
				list.addAll(orderlist);
			}
			rows = list.size();
			System.out.println("rows := " + rows);
			int pages = rows % 18 == 0 ? rows/18 : rows/18 + 1;
			System.out.println("pages := " + pages);
			if(pages > 1) {
				poi = new POIExcel();
				poi.specifyTemplate(inputStream, 1);
				
				String[] sheetNames = new String[pages];
				for(i = 0; i < pages; i ++) {
					sheetNames[i] = "" + (i + 1);
				}
				poi.cloneTemplate(1, sheetNames);
				poi.outputExcel(inputStream);
				poi.close();
			}
			for(i = 0; i < pages; i ++) {
				poi = new POIExcel();
				poi.specifyTemplate(inputStream, i + 1);
				poi.fillCellValue(3, 11, vehicle);
				poi.fillCellValue(2, 7, sim3.format(new Date()));
				poi.fillCellValue(1, 19, fac + "工厂");
				poi.fillCellValue(2, 2, username);
				poi.fillCellValue(2, 19, "焊装" + line +  "线");
				poi.fillCellValue(23, 21, version);
				poi.fillCellValue(26, 21, String.valueOf(pages));
				poi.fillCellValue(25, 21, String.valueOf(i + 1));
				for(j = 0; j < 18; j ++) {
					int dataindex = i * 18 + j;
					if(dataindex == rows) {
						break;
					}
					String[] rowdata = list.get(dataindex);
					int rowindex = j + 5;
					poi.fillCellValue(rowindex, 2, rowdata[0]);
					poi.fillCellValue(rowindex, 4, rowdata[1]);
					poi.fillCellValue(rowindex, 6, rowdata[2]);
					poi.fillCellValue(rowindex, 8, rowdata[3]);
					poi.fillCellValue(rowindex, 10, rowdata[4]);
					poi.fillCellValue(rowindex, 11, rowdata[5]);
					poi.fillCellValue(rowindex, 19, rowdata[6]);
					poi.fillCellValue(rowindex, 21, rowdata[7]);
				}
				if(i > 0) {
					poi.zoomSheet(1 + i, 26, 21, 50, true);
				}
				poi.outputExcel(inputStream);
				poi.close();
			}
			viewPanel.addInfomation("创建数据集，请耐心等待...\n", 90, 100);
			TCComponentDatasetType wordType = (TCComponentDatasetType) bopLine.getSession().getTypeComponent("MSExcelX");
			TCComponentDataset dataset = wordType.create(title, "", "MSExcelX");
			dataset.setFiles(new String[]{ inputStream }, new String[]{ "excel" });
			this.saveFiles(dataset);
			
			File file = new File(inputStream);
			file.delete();
			viewPanel.addInfomation("输出报表完成，请在选择保存的文件夹下查看！...\n", 100, 100);
			
//			OpenHomeDialog dialog = new OpenHomeDialog(AIFUtility.getActiveDesktop().getsh, session.getUser().getHomeFolder(),session);
//			dialog.open();
//			
//			datasetLocation = dialog.folder;
//			System.out.println("文件夹："+dialog.folder);
//			
//			if(dialog.flag) {
//				return ;
//			}
//			
//			if(datasetLocation == null ) {
//				return ;
//			}
		}catch(Exception e) {
			e.printStackTrace();
			MFCUtility.errorMassges("异常：" + e.getLocalizedMessage());
		}
	}
	private void getReportData(TCComponentBOMLine pline) {
		try {
			List<TCComponent> lstScope = new ArrayList<TCComponent>();
			if(this.virtualLines != null) {
				for(int i = 0; i < this.virtualLines.length; i ++) {
					lstScope.add(virtualLines[i]);
				}
			}else {
				lstScope.add(pline);
			}
//			List<TCComponent> tcclist = Util.callStructureSearch(lstScope, "__DFL_Find_Object_by_Name", new String[] { "NAME", "LX"},
//					new String[] { "*涂胶*", "B8_BIWArcWeldOP" });
//			System.out.println("涂胶工序：" + tcclist.size());
			List<TCComponent> tcclist1 = Util.callStructureSearch(lstScope, "__DFL_Find_Object_by_Name", new String[] { "NAME", "LX"},
					new String[] { "*涂胶*", "B8_BIWArcWeldOP" });
			List<TCComponent> tcclist2 = Util.callStructureSearch(lstScope, "__DFL_Find_Object_by_Name", new String[] { "NAME", "LX"},
					new String[] { "*弧焊*", "B8_BIWArcWeldOP" });
			System.out.println("涂胶工序：" + tcclist1.size());
			System.out.println("弧焊工序：" + tcclist2.size());
			List<TCComponent> tcclist = new ArrayList<TCComponent>();
			for(int i = 0; i < tcclist1.size() ; i ++) {
				TCComponentBOMLine line = (TCComponentBOMLine)tcclist1.get(i);
				System.out.println("涂胶：" + line.getItem().getProperty("item_id") + "/" + line.getItemRevision().getProperty("item_revision_id") + " := " + line);
				if(!tcclist.contains(tcclist1.get(i))) {
					tcclist.add(tcclist1.get(i));
				}
			}
			for(int i = 0; i < tcclist2.size() ; i ++) {
				TCComponentBOMLine line = (TCComponentBOMLine)tcclist2.get(i);
				System.out.println("弧焊：" + line.getItem().getProperty("item_id") + "/" + line.getItemRevision().getProperty("item_revision_id") + " := " + line);
				if(!tcclist.contains(tcclist2.get(i))) {
					tcclist.add(tcclist2.get(i));
				}
			}
			if(tcclist == null || tcclist.size() == 0) {
				return;
			}
			
			int i = 0, j = 0; 
			int count = tcclist.size();
			int cntWeld = 0;
			for(i = 0; i < count; i ++) {
				TCComponentBOMLine opLine = (TCComponentBOMLine)tcclist.get(i);
				if(!(opLine.getItem() instanceof TCComponentMEOP)) {
					System.out.println("涂胶：" + opLine + " 不是工序类型！");
					continue;
				}
				String opName = opLine.getItemRevision().getProperty("object_name");
				String body = this.getBodyinfo(opLine);
				if(body == null || body.length() == 0) {
					System.out.println("涂胶：" + opLine + " 工序未得到上层产线信息！");
					continue;
				}
				TcBOMService.expandOneLevel(session, new TCComponentBOMLine[] {opLine});
				AIFComponentContext[] children = opLine.getChildren();
				cntWeld = children.length;
				List<TCComponentItemRevision> lstWeld = new ArrayList<TCComponentItemRevision>();
				for(j = 0; j < cntWeld; j ++) {
					TCComponentBOMLine weldline = (TCComponentBOMLine)children[j].getComponent();
					if(weldline.getItem().getType().equals("ArcWeld")) {
						lstWeld.add(weldline.getItemRevision());
					}
				}
				if(lstWeld.size() > 0 ) {
					MFCUtility.loadProperties(session, lstWeld.toArray(new TCComponentItemRevision[0]), weldProps);
					cntWeld = lstWeld.size();
					for(j = 0; j < cntWeld; j ++) {
						TCComponentItemRevision weld = lstWeld.get(j);
						System.out.println(body + " --> weld is := " + weld);
						String[] rowdata = new String[8];
						rowdata[0] = "";
						rowdata[1] = "";
						rowdata[2] = "";
						String proportion = "";//this.hmTypeProportion.get(rowdata[1]);
						if(opName.contains("涂胶")) {
							rowdata[0] = "胶";
							rowdata[1] = weld.getPropertyDisplayableValue(weldProps[0]);
							if(!StringUtil.isEmpty(rowdata[1]) && this.hmTypeNes.containsKey(rowdata[1])) {
								rowdata[2] = this.hmTypeNes.get(rowdata[1]);
							}
							if(!StringUtil.isEmpty(rowdata[1]) && this.hmTypeProportion.containsKey(rowdata[1])) {
								proportion = this.hmTypeProportion.get(rowdata[1]);
							}
						}else if(opName.contains("弧焊")) {
							rowdata[0] = "焊丝";
							rowdata[2] = weld.getPropertyDisplayableValue(weldProps[0]);
							if(!StringUtil.isEmpty(rowdata[2]) && this.hmNesType.containsKey(rowdata[2])) {
								rowdata[1] = this.hmNesType.get(rowdata[2]) + "焊丝";
							}	
							if(!StringUtil.isEmpty(rowdata[2]) && this.hmNesProportion.containsKey(rowdata[2])) {
								proportion = this.hmNesProportion.get(rowdata[2]);
							}
						}
						System.out.println("DFL := " + rowdata[1]);
						System.out.println("NES := " + rowdata[2]);
						rowdata[3] = body;
						rowdata[4] = "g";
						rowdata[5] = "";
						String uom = weld.getPropertyDisplayableValue(weldProps[2]);
						
						String b8_Long = weld.getPropertyDisplayableValue(weldProps[1]);
						String b8_Diameter = weld.getPropertyDisplayableValue(weldProps[3]);
						String b8_Hight = weld.getPropertyDisplayableValue(weldProps[4]);
						System.out.println("直径b8_Diameter := " + b8_Diameter);
						System.out.println("比重rowdata[8] := " + proportion);
						System.out.println("长度b8_Long := " + b8_Long);
						System.out.println("高度b8_Hight := " + b8_Hight);
						if(proportion != null && proportion.length() > 0) {
							if(uom.equals("m") || uom.equals("米")) {
								try {
									rowdata[5] = new BigDecimal(3.1415).multiply(new BigDecimal(b8_Diameter))
											.multiply(new BigDecimal(b8_Diameter)).multiply(new BigDecimal(proportion))
											.multiply(new BigDecimal(b8_Long))
											.divide(new BigDecimal(4), 1, BigDecimal.ROUND_HALF_UP).toString();
									rowdata[5] = StringUtil.getStringCutZero(rowdata[5]);
								}catch(Exception e) {
									e.printStackTrace();
								}
							}else if(uom.equals("点")) {
								try {
									rowdata[5] = new BigDecimal(3.1415).multiply(new BigDecimal(b8_Diameter))
											.multiply(new BigDecimal(b8_Diameter)).multiply(new BigDecimal(proportion))
											.multiply(new BigDecimal(b8_Long))
											.multiply(new BigDecimal(b8_Hight))
											.divide(new BigDecimal(4000), 1, BigDecimal.ROUND_HALF_UP).toString();
									rowdata[5] = StringUtil.getStringCutZero(rowdata[5]);
								}catch(Exception e) {
									e.printStackTrace();
								}
							}
						}
						rowdata[6] = "市购";
						if(!StringUtil.isEmpty(rowdata[1])) {
							if(rowdata[1].startsWith("SD")) {
								rowdata[6] = "辉旭";
							}else if(rowdata[1].startsWith("ST") || rowdata[1].startsWith("#")) {
								rowdata[6] = "时利和";
							}
						}
						rowdata[7] = "包边班";
						if(mapBodyLineName.containsKey(body)) {
							String lineName = mapBodyLineName.get(body);
							if(lineName.contains("地板")) {
								rowdata[7] = "地板班";
							}else if(lineName.contains("侧围")) {
								rowdata[7] = "侧围班";
							}else if(lineName.contains("顶盖")) {
								rowdata[7] = "顶盖班";
							}else if(lineName.contains("主线")) {
								rowdata[7] = "主线班";
							}else if(lineName.contains("机舱")) {
								rowdata[7] = "机舱班";
							}else {
								rowdata[7] = "包边班";
							}
						}
						
						if(this.hmBodyData.containsKey(body)) {
							List<String[]> list = this.hmBodyData.get(body);
							list.add(rowdata);
							rows ++;
							this.hmBodyData.put(body, list);
						}else {
							List<String[]> list = new ArrayList<String[]>();
							list.add(rowdata);
							rows ++;
							this.hmBodyData.put(body, list);
						}
						if(!this.lstBodies.contains(body)) {
							this.lstBodies.add(body);
						}
					}
				}
			}
		}catch(Exception e) {
			e.printStackTrace();
		}
	}
	private String getBodyinfo(TCComponentBOMLine opLine) {
		String body = "";
		try {
			TCComponentBOMLine statLine = opLine.parent();
			if(this.hmLineBody.containsKey(statLine)) {
				return this.hmLineBody.get(statLine);
			}
			TCComponentBOMLine lineLine = statLine.parent();
			TcBOMService.expandOneLevel(session, new TCComponentBOMLine[] {lineLine});
			AIFComponentContext[] children = lineLine.getChildren();
			int i = 0;
			int count = children.length;
			int cntStat = 0;
			for(i = 0; i < count; i ++) {
				TCComponentBOMLine sline = (TCComponentBOMLine)children[i].getComponent();
				if(sline.getItem().getType().equals("B8_BIWMEProcStat")) {
					cntStat ++;
					if(cntStat > 1) {
						break;
					}
				}
			}
			if(cntStat == 1) {
				body = lineLine.getItemRevision().getProperty("b8_ChineseName");
				if(!StringUtil.isEmpty(body)) {
					mapBodyLineName.put(body, body);
					System.out.println("body  " + body + " --> chinesename := " + body);
				}
				if(body.length() == 0) {
					body = "-";
				}
				this.hmLineBody.put(lineLine, body);
				this.hmLineBody.put(statLine, body);
//				if(!this.lstBodies.contains(body)) {
//					this.lstBodies.add(body);
//				}
			}else if(cntStat > 1) {
				body = lineLine.getItemRevision().getProperty("b8_ChineseName") + "\n" + statLine.getItemRevision().getProperty("object_name");
				if(!StringUtil.isEmpty(lineLine.getItemRevision().getProperty("b8_ChineseName"))) {
					mapBodyLineName.put(body, lineLine.getItemRevision().getProperty("b8_ChineseName"));
					System.out.println("body  " + body + " --> chinesename := " + lineLine.getItemRevision().getProperty("b8_ChineseName"));
				}
				this.hmLineBody.put(statLine, body);
//				if(!this.lstBodies.contains(body)) {
//					this.lstBodies.add(body);
//				}
			}
		} catch (Exception e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		
		return body;
	}
	public void saveFiles(TCComponentDataset ds) {
		try {
				int i = 0;
				Map<String, Object> itemMap = new HashMap<String, Object>();
				Map<String, Object> itemRevisionMap = new HashMap<String, Object>();
				Map<String, Object> itemRevMasterFormMap = new HashMap<String, Object>();
				itemMap.put("item_id", ""); //$NON-NLS-1$ //$NON-NLS-2$
				itemMap.put("object_name", title); //$NON-NLS-1$
				itemMap.put("object_desc", ""); //$NON-NLS-1$
				itemMap.put("object_type", "DFL9MEDocument"); //$NON-NLS-1$
				itemRevisionMap.put("object_type", "DFL9MEDocumentRevision"); //$NON-NLS-1$
				itemRevisionMap.put("object_name", title); //$NON-NLS-1$
				itemRevisionMap.put("dfl9_process_type", "H"); //$NON-NLS-1$
				itemRevisionMap.put("dfl9_process_file_type", "ZC"); //$NON-NLS-1$
				//itemRevisionMap.put("dfl9_vehiclePlant", "docNo"); 
				itemRevMasterFormMap.put("object_type", "DFL9MEDocumentRevisionMaster"); //$NON-NLS-1$
				CreateResponse respose = TCComponentUtils.create(itemMap, itemRevisionMap, itemRevMasterFormMap);
				int num = respose.serviceData.sizeOfCreatedObjects();
				TCComponentItemRevision rev = null;
				TCComponentItem tccomponentitem = null;
				if(num > 0){
					for(i=0;i<num;i++){
						TCComponent comp = respose.serviceData.getCreatedObject(i);
						if(comp instanceof TCComponentItemRevision){
							rev = (TCComponentItemRevision) comp;						
						}else if(comp instanceof TCComponentItem) {
							tccomponentitem = (TCComponentItem)comp;
						}
					}
				}
				// 添加文档与数据集的关系
				rev.add("IMAN_specification", ds);
				// 添加焊装工位与文档的关系
				if(datasetLocation instanceof TCComponentFolder) {
					datasetLocation.add("contents", tccomponentitem);
				}
		} catch (Exception e) {
			e.printStackTrace();
		}
	}
}
