package com.dfl.report.mfcadd;

import java.io.File;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Collections;
import java.util.Comparator;
import java.util.Date;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import com.dfl.report.util.ReportViwePanel;
import com.dfl.report.util.Util;
import com.teamcenter.rac.aif.kernel.AIFComponentContext;
import com.teamcenter.rac.kernel.TCComponent;
import com.teamcenter.rac.kernel.TCComponentBOMLine;
import com.teamcenter.rac.kernel.TCComponentBOMWindow;
import com.teamcenter.rac.kernel.TCComponentDataset;
import com.teamcenter.rac.kernel.TCComponentDatasetType;
import com.teamcenter.rac.kernel.TCComponentFolder;
import com.teamcenter.rac.kernel.TCComponentItemRevision;
import com.teamcenter.rac.kernel.TCException;
import com.teamcenter.rac.kernel.TCSession;
import com.teamcenter.rac.kernel.variants.BOMVariantOptionModel;
import com.teamcenter.rac.kernel.variants.BOMVariantOptionValueModel;
import com.teamcenter.rac.kernel.variants.BOMVariantRuleModel;
import com.teamcenter.rac.kernel.variants.BOMVariantRulePartialException;
import com.teamcenter.rac.kernel.variants.BOMVariantSOAHelper;

public class VehiclePSummaryReportOperation {
TCComponentBOMLine bopLine = null;
TCComponent datasetLocation = null;
String title = "";
String curdate = "";
private SimpleDateFormat dateFormat = new SimpleDateFormat("yyyy.M.dd");
int rows = 0;
private List<TCComponentBOMLine> lstMoveLine;
private TCComponentBOMLine ebomline = null;
private HashMap<TCComponentBOMLine, String> hmMovePartition;
private List<String> lstArea;
private HashMap<String, List<String>> hmAreaFuncs;
private HashMap<String, List<String>> hmAFPartNums;
private HashMap<String, String[]> hmAFParts;
//private HashMap<TCComponentBOMLine, TCComponentBOMLine> hmLineFuncs;
//private HashMap<TCComponentBOMLine, TCComponentBOMLine> hmLineFALine;
private HashMap<TCComponentBOMLine, String[]> hmFuncInfo;
private List<String> lstVechil = null;
private static final String bl_DFL9SolItmPartRevision_dfl9_part_no = "bl_DFL9SolItmPartRevision_dfl9_part_no";
private TCSession session = null;
//private List<String> lstSamePartNo = null;
TCComponentBOMWindow bomwindow = null;
private static String regex = "^(\\d{2}-)[^']*";
private List<String> listBBOMID = null;
private List<String> listReportKey = null;
private final String KEYINFO = "@@@";
private Map<String, String> mapPartVariant;//mifc 20200313
public VehiclePSummaryReportOperation(TCComponentBOMLine bop, TCComponentBOMLine ebom, TCComponentBOMWindow window, TCComponent folder) {
	bopLine = bop;
	session = bopLine.getSession();
	ebomline = ebom;
	datasetLocation = folder;
	lstMoveLine = new ArrayList<TCComponentBOMLine>();
	hmMovePartition = new HashMap<TCComponentBOMLine, String>();
//	hmLineFuncs = new HashMap<TCComponentBOMLine, TCComponentBOMLine>();
//	hmLineFALine = new HashMap<TCComponentBOMLine, TCComponentBOMLine>();
	lstArea = new ArrayList<String>();
	hmAreaFuncs = new HashMap<String, List<String>>();
	hmAFPartNums = new HashMap<String, List<String>>();
	hmAFParts = new HashMap<String, String[]>();
	hmFuncInfo = new HashMap<TCComponentBOMLine, String[]>();
	lstVechil = new ArrayList<String>();
//	lstSamePartNo = new ArrayList<String>();
	bomwindow = window;
	listBBOMID = new ArrayList<String>();
	listReportKey = new ArrayList<String>();
	mapPartVariant = new HashMap<String, String>();
	getVariantOptionValues();
	getAndoutReport();
	
}
private void getVariantOptionValues() {
	try {
		List<BOMVariantRuleModel> list = BOMVariantSOAHelper.getVariantRules(ebomline.window());
		System.out.println("getVariantRules list.size := " + list.size());
		int i = 0;
		int size = list.size();
		BOMVariantRuleModel ruleModel;
		BOMVariantOptionModel optionModel;
		List<BOMVariantOptionModel> listOptions;
		int j = 0, k = 0;
		int count = 0, length = 0;
		String optionName;
		List<BOMVariantOptionValueModel> listOpValues;
		BOMVariantOptionValueModel optionValue;
		for(i = 0; i < size; i ++) {
			ruleModel = list.get(i);
			listOptions = ruleModel.getOptions();
			count = listOptions.size();
			for(j = 0;j < count; j ++) {
				optionModel = listOptions.get(j);
				optionName = optionModel.getOptionName();
				System.out.println("optionName := " + optionName);
				listOpValues = optionModel.getOptionValues();
				length = listOpValues.size();
				if(!optionName.equals("veh")) {
					continue;
				}
				for(k = 0; k < length; k ++) {
					optionValue = listOpValues.get(k);
					System.out.println("optionValue := " + optionValue.getValue());
					this.lstVechil.add(optionValue.getValue());
				}
			}
		}
	} catch (BOMVariantRulePartialException e) {
		// TODO Auto-generated catch block
		e.printStackTrace();
	} catch (IllegalArgumentException e) {
		// TODO Auto-generated catch block
		e.printStackTrace();
	} catch (TCException e) {
		// TODO Auto-generated catch block
		e.printStackTrace();
	}
}
public void getAndoutReport() {
	try {
		// 界面显示进度并输出执行步骤
		ReportViwePanel viewPanel = new ReportViwePanel("生成报表");
		viewPanel.setVisible(true);
		viewPanel.addInfomation("正在获取模板...\n", 20, 100);
		// 查询并导出模板
		String inputStream = TemplateUtil.getTemplateFile("DFL_Template_VehicleModleSummary");
		if (inputStream == null) {
			viewPanel.addInfomation("错误：没有找到车型零件式样差信息汇总表的模板，请先在TC中添加模板(名称为：DFL_Template_VehicleModleSummary)\n", 100,100);
			bomwindow.close();
			bomwindow = null;
			return;
		}
		viewPanel.addInfomation("开始输出报表...\n", 35, 100);
		String familycode = bopLine.getItemRevision().getProperty("project_ids");// 车型
		String vehicle = Util.getDFLProjectIdVehicle(familycode);
		String phase = "";
		String bopName = bopLine.getItemRevision().getProperty("object_name");
		String[] splits = bopName.split("_");
		if(splits.length > 3) {
			phase = splits[3];
		}
		title = vehicle + "车型零件式样差信息汇总表（" + phase + "阶段）";
		SimpleDateFormat sim = new SimpleDateFormat("yyyy.M.dd");
		curdate = sim.format(new Date());
		getReportData(this.bopLine);
		viewPanel.addInfomation("开始写数据，请耐心等待...\n", 70, 100);
		POIExcel poi = new POIExcel();
		poi.specifyTemplate(inputStream);
		//String[] vechils = MFCUtility.getVariantValues(this.ebomline);
		
		int newCols = this.lstVechil.size();
//		if(vechils != null && vechils.length > newCols) {
//			this.lstVechil.clear();
//			for(int i = 0; i < vechils.length; i ++) {
//				this.lstVechil.add(vechils[i]);
//			}
//			newCols = this.lstVechil.size();
//		}
		Collections.sort(lstVechil);
		System.out.println("newCols := " + newCols);
		if(newCols > 1) {
			poi.copyCell(0, 7, 0, 8, 6 + newCols, true);
			poi.copyCell(1, 7, 1, 8, 6 + newCols, true);
			poi.copyCell(2, 7, 2, 8, 6 + newCols, true);
			for(int row = 3; row < 11; row ++) {
				poi.copyCell(3, 7, row, 8, 6 + newCols, true);
			}
			poi.addMergedRegion(0, 7, 0, 6 + newCols);
			poi.addMergedRegion(1, 7, 1, 6 + newCols);
			for(int col = 0; col < newCols; col ++) {
				poi.fillCellValue(2, 7 + col, this.lstVechil.get(col));
			}
		}else if(newCols == 1) {
			poi.fillCellValue(2, 7, this.lstVechil.get(0));
		}
		poi.fillCellValue(0, 7, "日期：" + curdate);
		poi.fillCellValue(1, 7, "整车车型");
		poi.fillCellValue(0, 0, title);
		rows = this.hmAFParts.size();
		System.out.println("rows := " + rows);
		if(rows > 8) {
			poi.appendRow(10, rows - 8);
		}
		int i = 0, j = 0, k = 0;
		int cntArea = 0, cntFuncs = 0, cntParnos = 0;
		List<String> lstFuncs, lstPartnos;
		String area = "", function = "", funcname = "", partno = "";
		cntArea = this.lstArea.size();
		System.out.println("cntArea := " + cntArea);
		int rowIndex = 3, funcIndex = 3, areaIndex = 3;
		for(i = 0; i < cntArea; i ++) {
			area = this.lstArea.get(i);
			System.out.println("area := " + area);
			lstFuncs = this.hmAreaFuncs.get(area);
			cntFuncs = lstFuncs.size();
			for(j = 0; j < cntFuncs; j ++) {
				function = lstFuncs.get(j);
				System.out.println("function := " + function);
				lstPartnos = this.hmAFPartNums.get(area + KEYINFO + function);
				
				Comparator comparator = new Comparator() {
					public int compare(Object paramObject1, Object paramObject2) {
						String partno1 = (String) paramObject1;
						String partno2 = (String) paramObject2;
						return partno1.compareTo(partno2);
					}
				};
				Collections.sort(lstPartnos,comparator);
				
				cntParnos = lstPartnos.size();
				for(k = 0; k < cntParnos; k ++) {
					partno = lstPartnos.get(k);
					System.out.println("partno := " + partno);
					String[] infos = this.hmAFParts.get(area + KEYINFO + function + KEYINFO + partno);
					if(infos == null || infos.length == 0){
						continue;
					}
					funcname = infos[1];
					poi.fillCellValue(rowIndex, 0, String.valueOf(rowIndex - 2));
					poi.fillCellValue(rowIndex, 4, infos[2]);
					poi.fillCellValue(rowIndex, 5, infos[3]);
					poi.fillCellValue(rowIndex, 6, infos[4]);
					String dfl_partno = partno.split(KEYINFO)[0];
					String varaint = this.mapPartVariant.get(dfl_partno);//mifc 20200313
					System.out.println("dfl_partno := " + dfl_partno + " --> varaint := " + varaint);
					//if(!StringUtil.isEmpty(infos[5])) {
					if(!StringUtil.isEmpty(varaint)) {//mifc 20200313
						for(int col = 0; col < newCols; col ++) {
							if(varaint.contains(this.lstVechil.get(col))) {//mifc 20200313
								poi.fillCellValue(rowIndex, 7 + col, "〇");
							}
						}
					}
					poi.fillCellValue(rowIndex, 2, function);
					poi.fillCellValue(rowIndex, 3, funcname);
					rowIndex ++;
				}
//				if((rowIndex - 1) > funcIndex) {
//					poi.addMergedRegion(funcIndex, 2, rowIndex - 1, 2);
//					poi.addMergedRegion(funcIndex, 3, rowIndex - 1, 3);
//				}
//				poi.fillCellValue(funcIndex, 2, function);
//				poi.fillCellValue(funcIndex, 3, funcname);
				funcIndex = rowIndex;
			}
			if((rowIndex - 1) > areaIndex) {
				poi.addMergedRegion(areaIndex, 1, rowIndex - 1, 1);
			}
			poi.fillCellValue(areaIndex, 1, area);
			areaIndex = rowIndex;
		}
		SimpleDateFormat sim2  = new SimpleDateFormat("yyyyMMdd_HH"	);
		title = title + "_" + sim2.format(new Date());
		String newName = System.getenv("TMP") + File.separator + MFCUtility.fileNameReplace(title, "") + ".xlsx"; 
		poi.outputExcel(newName);
		File ftmp = new File(inputStream);
		ftmp.delete();
		inputStream = newName;
		viewPanel.addInfomation("创建数据集，请耐心等待...\n", 90, 100);
		TCComponentDatasetType wordType = (TCComponentDatasetType) bopLine.getSession().getTypeComponent("MSExcelX");
		
		TCComponentDataset dataset = wordType.create(title, "", "MSExcelX");
		dataset.setFiles(new String[]{ inputStream }, new String[]{ "excel" });
		if(datasetLocation instanceof TCComponentFolder) {
			datasetLocation.add("contents", dataset);
		}else if(datasetLocation instanceof TCComponentItemRevision) {
			datasetLocation.add("IMAN_specification", dataset);
		}
		
		File file = new File(inputStream);
		file.delete();
		viewPanel.addInfomation("输出报表完成，请在选择保存的文件夹下查看！...\n", 100, 100);
	}catch(Exception e) {
		e.printStackTrace();
		MFCUtility.errorMassges("异常：" + e.getLocalizedMessage());
	}finally {
		if(bomwindow != null) {
			try {
				bomwindow.close();
			} catch (TCException e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			}
			bomwindow = null;
		}
		
	}
}
private TCComponentBOMLine functionLine = null;
private TCComponentBOMLine taLine = null;
private List<TCComponentBOMLine> listTalines ;
private List<String> listKeys;
private void getReportData(TCComponentBOMLine pline) {
	try {
		List tcclist = Util.callStructureSearch(pline, "__DFL_Find_SolutionPart", new String[] { "PARTNO" },
				new String[] { "*" });
		System.out.println("tcclist.size := " + tcclist.size());
		int i = 0;
		int size = tcclist.size();
		//在查询的结果上，找到移动单元，及移动单元对应的分区
		MFCUtility.loadProperties(session, (TCComponentBOMLine[])tcclist.toArray(new TCComponentBOMLine[0]), new String[] {"bl_usage_address", "bl_item_item_id", bl_DFL9SolItmPartRevision_dfl9_part_no});
		for(i = 0; i < size; i ++) {
			TCComponentBOMLine solLine = (TCComponentBOMLine) tcclist.get(i);
			String bl_usage_address = solLine.getPropertyDisplayableValue("bl_usage_address");
			String itemid = solLine.getPropertyDisplayableValue("bl_item_item_id");
			if(!this.listBBOMID.contains(itemid)) {
				this.listBBOMID.add(itemid);
			}
			//System.out.println("bl_usage_address := " + bl_usage_address);
			if(bl_usage_address.length() > 0 && bl_usage_address.startsWith("MU")) {
				this.lstMoveLine.add(solLine);
				this.getPartition(solLine, solLine);
			}
		}
		TCSession session = pline.getSession();
		MFCUtility.loadProperties(session, this.lstMoveLine.toArray(new TCComponentBOMLine[0]), new String[] {bl_DFL9SolItmPartRevision_dfl9_part_no});
		//到ebom里搜索ta
		size = this.lstMoveLine.size();
		System.out.println("BBOM移动单元 := " + size);
		int j = 0, k = 0, m = 0;
		int count = 0, cntFuncs = 0, cntTas = 0;
		List<String> listPartnos = new ArrayList<String>();
		Map<String, TCComponentBOMLine[]> mapPartSols = new HashMap<String, TCComponentBOMLine[]>();
		Map<TCComponentBOMLine, TCComponentBOMLine> mapSolFunction = new HashMap<TCComponentBOMLine, TCComponentBOMLine>();
		Map<TCComponentBOMLine, TCComponentBOMLine> mapSolTaline = new HashMap<TCComponentBOMLine, TCComponentBOMLine>();
		Map<String, List<TCComponentBOMLine>> mapSolFunctions = new HashMap<String, List<TCComponentBOMLine>>();
		List<TCComponentBOMLine> listFunctions;
		Map<TCComponentBOMLine, String> mapPartnoKeys = new HashMap<TCComponentBOMLine, String>();
		TCComponentBOMLine function, fromTA;
		String partKey = "";
		for(i = 0; i < size; i ++) {//遍历BBOM中的移动单元
			TCComponentBOMLine solLine = lstMoveLine.get(i);
			String area = this.hmMovePartition.get(solLine);
			if(area == null) {
				area = "";
			}
			System.out.println("area := " + area);
			String dfl9_part_no = solLine.getPropertyDisplayableValue(bl_DFL9SolItmPartRevision_dfl9_part_no);//solLine.getItemRevision().getProperty("dfl9_part_no");
			System.out.println("dfl9_part_no := " + dfl9_part_no);
			if(!StringUtil.isEmpty(dfl9_part_no)) {//获取图号，并到EBOM中查询相同图号的EBOMLIne行
				if(dfl9_part_no.matches(regex)) {
					System.out.println("dfl9_part_no := " + dfl9_part_no + "是标准件");
					continue;
				}
				if(mapPartVariant.containsKey(dfl9_part_no)) {//mifc 20200313
					
					String newvaraint = MFCUtility.getVariantConditions(solLine);
					if(!StringUtil.isEmpty(newvaraint)) {
						String oldvaraint = mapPartVariant.get(dfl9_part_no);
						List<String> listVars = new ArrayList<String>();
						String[] newvaraints = newvaraint.split(",");
						String[] oldvaraints = oldvaraint.split(",");
						StringBuffer sbNew = new StringBuffer();
						for(j = 0; j < newvaraints.length; j ++) {
							if(listVars.contains(newvaraints[j])) {
								continue;
							}
							listVars.add(newvaraints[j]);
							if(sbNew.toString().length() >0) {
								sbNew.append(",");
							}
							sbNew.append(newvaraints[j]);
						}
						for(j = 0; j < oldvaraints.length; j ++) {
							if(listVars.contains(oldvaraints[j])) {
								continue;
							}
							listVars.add(oldvaraints[j]);
							if(sbNew.toString().length() >0) {
								sbNew.append(",");
							}
							sbNew.append(oldvaraints[j]);
						}
						mapPartVariant.put(dfl9_part_no, sbNew.toString());
					}
				}else {
					String varaint = MFCUtility.getVariantConditions(solLine);
					if(!StringUtil.isEmpty(varaint)) {
						mapPartVariant.put(dfl9_part_no, varaint);
					}
				}
				if(listPartnos.contains(area + KEYINFO + dfl9_part_no)) {//不同分区的图号不同
					continue;//第1重过滤，过滤掉BBOM相同分区下移动单元的零件号相同的查询项
				}
				listPartnos.add(area + KEYINFO + dfl9_part_no);
				TCComponentBOMLine[] ebomLines = null;
				if(mapPartSols.containsKey(dfl9_part_no)) {
					ebomLines = mapPartSols.get(dfl9_part_no);
					count = ebomLines.length;
					System.out.println("已查询过的：dfl9_part_no := " + dfl9_part_no + " --> sols count := " + ebomLines.length );
				}else {
					List ebomSols = Util.callStructureSearch(this.ebomline, "__DFL_Find_SolutionPart", new String[] { "PARTNO" },
							new String[] {dfl9_part_no});
					count = ebomSols.size();
					System.out.println("dfl9_part_no := " + dfl9_part_no + " --> sols count := " + count );
					if(count < 1) {
						mapPartSols.put(dfl9_part_no, new TCComponentBOMLine[0]);
						continue;
					}
					ebomLines = new TCComponentBOMLine[count];//(TCComponentBOMLine[])ebomSols.toArray(new Object[0]);
					for(j = 0; j < count; j ++) {
						ebomLines[j] = (TCComponentBOMLine)ebomSols.get(j);
					}
					mapPartSols.put(dfl9_part_no, ebomLines);
				}
				if(ebomLines == null || ebomLines.length == 0) {
					continue;
				}
				if(mapSolFunctions.containsKey(dfl9_part_no)) {
					listFunctions = mapSolFunctions.get(dfl9_part_no);
				}else {//获取function
					listFunctions = new ArrayList<TCComponentBOMLine>();
					for(j = 0; j < count; j ++) {
						if(mapSolFunction.containsKey(ebomLines[j])) {
							function = mapSolFunction.get(ebomLines[j]);
							fromTA = mapSolTaline.get(ebomLines[j]);
							if(!listFunctions.contains(function)) {
								listFunctions.add(function);
							}
						}else {
							this.functionLine = null;
							this.taLine = null;
							this.getFuncs(ebomLines[j]);
							function = this.functionLine;
							fromTA = this.taLine;
							mapSolFunction.put(ebomLines[j], function);
							mapSolTaline.put(ebomLines[j], fromTA);
							if(function != null && !listFunctions.contains(function)) {
								listFunctions.add(function);
							}
						}
					}
					mapSolFunctions.put(dfl9_part_no, listFunctions);
				}
				if(listFunctions ==  null || listFunctions.size() == 0) {
					System.out.println("未获取到对应的function，无法输出数据。。。。");
					continue;
				}
				System.out.println("listFunctions.size ;= " + listFunctions.size());
				MFCUtility.loadProperties(session, ebomLines, new String[] {bl_DFL9SolItmPartRevision_dfl9_part_no, "DFL9_from_ta", "B8_MFGRemark", "bl_condition_tag"});
				//List<TCComponentBOMLine> lstFunctions = new ArrayList<TCComponentBOMLine>();//记录已经遍历过的ta，原来是记录function，后来改掉20200210
				cntFuncs = listFunctions.size();
				listTalines = new ArrayList<TCComponentBOMLine>();//相同号，记录已经输出过的ta
				listKeys = new ArrayList<String>();
				String funcno = "";
				String funcname = "";
				String vehVaraint = "";
				String remark = "";
				String varaint = "";
				for(j = 0; j < count; j ++) {
					function = mapSolFunction.get(ebomLines[j]);
					fromTA = mapSolTaline.get(ebomLines[j]);
					if(function == null) {
						System.out.println(ebomLines[j] + " 未找到function");
						continue;
					}
					if(fromTA == null) {
						System.out.println(ebomLines[j] + " 未找到taline");
						continue;
					}
					if(mapPartnoKeys.containsKey(ebomLines[j])) {
						partKey = mapPartnoKeys.get(ebomLines[j]);
						String[] splits = partKey.split(KEYINFO);
						funcno = splits[0];
						funcname = splits[1];
						vehVaraint = splits[2];
						remark = splits[3];
						varaint = splits[4];
						if(vehVaraint.equals("___")) {
							vehVaraint = "";
						}
						if(remark.equals("___")) {
							remark = "";
						}
						if(varaint.equals("___")) {
							varaint = "";
						}
					}else {
						if(this.hmFuncInfo.containsKey(function)) {
							String[] info = this.hmFuncInfo.get(function);
							funcno = info[0];
							funcname = info[1];
						}else {
							String[] info = new String[2];
							info[0] = function.getPropertyDisplayableValue("bl_rev_object_name");
							info[1] = function.getPropertyDisplayableValue("bl_rev_object_desc");
							funcno = info[0];
							funcname = info[1];
							hmFuncInfo.put(function, info);
						}
						String firstFA = fromTA.getPropertyDisplayableValue("DFL9_from_ta");
						System.out.println("DFL9_from_ta := " + firstFA);
						if(firstFA == null || firstFA.length() == 0) {
							System.out.println("DFL9_from_ta is null,即式样为空，不再读取数据");
							continue;
						}
						String[] fromtas = firstFA.split("_");
						vehVaraint = firstFA;
						if(fromtas.length >= 3) {
							vehVaraint = fromtas[2];
						}else {
							System.out.println("DFL9_from_ta 的值不符合标准,即式样值无法拆分，不再读取数据");
							continue;
						}
						remark = ebomLines[j].getPropertyDisplayableValue("B8_MFGRemark");
						varaint = MFCUtility.getVariantConditions(ebomLines[j]);
						StringBuffer sbKey = new StringBuffer();
						sbKey.append(funcno).append(KEYINFO);
						sbKey.append(funcname).append(KEYINFO);
						sbKey.append(vehVaraint).append(KEYINFO);
						if(StringUtil.isEmpty(remark)) {
							sbKey.append("___").append(KEYINFO);
						}else {
							sbKey.append(remark).append(KEYINFO);
						}
						if(StringUtil.isEmpty(varaint)) {
							sbKey.append("___");
						}else {
							sbKey.append(varaint);
						}
						partKey = sbKey.toString();
						mapPartnoKeys.put(ebomLines[j], sbKey.toString());
					}
					System.out.println("partKey := " + partKey);
					if(this.listKeys.contains(partKey)) {
						continue;//第2层，过滤掉同一个function下ebom行信息相同的查询项
					}
					this.listKeys.add(partKey);
					for(k = 0; k < cntFuncs; k ++) {
						function = listFunctions.get(k);
						AIFComponentContext[] tas = function.getChildren();
						cntTas = tas.length;
						for(m = 0; m < cntTas; m ++) {
							fromTA = (TCComponentBOMLine)tas[m].getComponent();
							
							String firstFA = fromTA.getPropertyDisplayableValue("DFL9_from_ta");
							System.out.println("DFL9_from_ta := " + firstFA);
							if(firstFA == null || firstFA.length() == 0) {
								System.out.println("DFL9_from_ta is null,即式样为空，不再读取数据");
								continue;
							}
							String[] fromtas = firstFA.split("_");
							String cxbl = firstFA;
							if(fromtas.length >= 3) {
								cxbl = fromtas[2];
							}else {
								System.out.println("DFL9_from_ta 的值不符合标准,即式样值无法拆分，不再读取数据");
								continue;
							}
							System.out.println("cxbl := " + cxbl);
							if(cxbl.equals(vehVaraint)) {
								continue;
							}
							if(listTalines.contains(fromTA)) {
								continue;//已经输出过的ta
							}
							listTalines.add(fromTA);
							gatherTheReportData(ebomLines[j], area, vehVaraint, dfl9_part_no, cxbl, remark, fromTA, function, true);
						}
					}
				}
				
//				for(j = 0; j < count; j ++) {
//					this.functionLine = null;
//					this.taLine = null;
//					this.getFuncs(ebomLines[j]);
//					if(this.functionLine == null) {
//						System.out.println(ebomLines[j] + " 未找到function");
//						continue;
//					}
//					System.out.println( "function := " + functionLine);
//					System.out.println( "taLine := " + taLine);
//					funcLine = this.taLine;
////					if(lstFunctions.contains(funcLine)) {
////						System.out.println("相同的ta，即与之前的式样相同的，则不再去查找");
////						continue;
////					}
////					lstFunctions.add(funcLine);
//					String firstFA = taLine.getPropertyDisplayableValue("DFL9_from_ta");
//					System.out.println("DFL9_from_ta := " + firstFA);
//					if(firstFA == null || firstFA.length() == 0) {
//						System.out.println("DFL9_from_ta is null,即式样为空，不再读取数据");
//						continue;
//					}
//					String[] fromtas = firstFA.split("_");
//					String firstCXBL = firstFA;
//					if(fromtas.length >= 3) {
//						firstCXBL = fromtas[2];
//					}else {
//						System.out.println("DFL9_from_ta 的值不符合标准,即式样值无法拆分，不再读取数据");
//						continue;
//					}
//					String remark = ebomLines[j].getPropertyDisplayableValue("B8_MFGRemark");
//					String varaintRule = MFCUtility.getVariantConditions(ebomLines[j]);
//					String key = firstCXBL + KEYINFO + remark + KEYINFO +  varaintRule;
//					System.out.println("key := " + key);
//					if(listKeys.contains(key)) {
//						continue;
//					}
//					getTheCheckPart(area, dfl9_part_no, firstCXBL, remark, functionLine , ebomLines[j]) ;//DFL9_from_ta - cxbl把参与比较的bom写进去
//					getFunctionDatas(ebomLines[j],area, firstCXBL, dfl9_part_no, remark);
//				}
//				List<TCComponentBOMLine> lstFALines = new ArrayList<TCComponentBOMLine>();//TA
//				List<TCComponentBOMLine> lstFuncLines = new ArrayList<TCComponentBOMLine>();
//				for(j = 0; j < count; j ++) {//遍历获取ebom行对应的FA（上层就是function）及function
//					TCComponentBOMLine ebom = ebomLines[j];//(TCComponentBOMLine)ebomSols.get(j);
//					this.getFuncs(ebom, ebom, lstFALines, lstFuncLines);//一定会得到FA、function吗
//				}
//				TCComponentBOMLine[] faLines = lstFALines.toArray(new TCComponentBOMLine[0]);
//				TCComponentBOMLine[] fuLines = lstFuncLines.toArray(new TCComponentBOMLine[0]);
//				MFCUtility.loadProperties(session, faLines, new String[] {"DFL9_from_ta", "bl_sequnce_no"});
//				MFCUtility.loadProperties(session, fuLines, new String[] {"bl_rev_object_name", "bl_rev_object_desc"});
//				int index = getFirstFA(ebomLines);//出现一个子层一个顶层，判断式样是否存在差，同式样不输出，不同式样在子级所在一级中查找进行输出
//																	//出现搜索结果均为同一式样顶层，不进行判断，直接结束；不同式样，在第二个式样下面找
//				if(!this.hmLineFALine.containsKey(ebomLines[index])) {
//					System.out.println("未得到FA的EBOM：" + ebomLines[index]);
//					continue;
//				}
//				String firstFA = this.hmLineFALine.get(ebomLines[index]).getPropertyDisplayableValue("DFL9_from_ta");
//				TCComponentBOMLine firstFunLine = this.hmLineFuncs.get(ebomLines[index]);
//				String[] fromtas = firstFA.split("_");
//				String firstCXBL = firstFA;
//				if(fromtas.length > 3) {
//					firstCXBL = fromtas[2];
//				}
//				String remark = ebomLines[index].getPropertyDisplayableValue("B8_MFGRemark");
//				for(j = 0; j < count; j ++) {
////					if(j == index) {
////						continue;
////					}
//					if(!this.hmLineFALine.containsKey(ebomLines[j])) {
//						System.out.println("未得到FA的EBOM：" + ebomLines[j]);
//						continue;
//					}
//					if(!this.hmLineFuncs.containsKey(ebomLines[j])) {
//						System.out.println("未得到Function的EBOM：" + ebomLines[j]);
//						continue;
//					}
//					TCComponentBOMLine faLine = this.hmLineFALine.get(ebomLines[j]);
//					String DFL9_from_ta = faLine.getPropertyDisplayableValue("DFL9_from_ta");
//					if(DFL9_from_ta.equals(firstFA)) {//出现搜索结果均为同一式样顶层，不进行判断，直接结束；
//						continue;
//					}
//					String[] from_tas = DFL9_from_ta.split("_");
//					String cxbl = DFL9_from_ta;
//					if(from_tas.length > 3) {
//						cxbl = from_tas[2];
//					}
//					if(faLine == ebomLines[j]) {
//						gatherTheReportData(ebomLines[j], area, firstCXBL, dfl9_part_no, cxbl, remark, 
//								ebomLines[j] , firstFunLine, true) ;
//					}else {
//						gatherTheReportData(ebomLines[j], area, firstCXBL, dfl9_part_no, cxbl,remark, 
//								ebomLines[j].parent() , firstFunLine, false) ;
//					}
//				}
			}
		}
	}catch(Exception e) {
		e.printStackTrace();
	}
}
private void getFunctionDatas(TCComponentBOMLine ebomLine, String area, String firstCXBL, String dfl9_part_no, String remark) {
	try {
		AIFComponentContext[] tas = this.functionLine.getChildren();
		int i = 0;
		int count = tas.length;
		TCComponentBOMLine taline = null;
		System.out.println(ebomLine + " --> taLine := " + this.taLine + " --> firstCXBL := " + firstCXBL); 
		for(i = 0; i < count; i ++) {
			taline = (TCComponentBOMLine)tas[i].getComponent();
			if(taline == this.taLine) {
				System.out.println("同一个taLine ：= " + i);
				continue;
			}
			if(listTalines.contains(taline)) {
				continue;//已经输出过的ta
			}
			listTalines.add(taline);
			String DFL9_from_ta = taline.getPropertyDisplayableValue("DFL9_from_ta");
			System.out.println(taline + " --> DFL9_from_ta := " + DFL9_from_ta);
			if(DFL9_from_ta == null || DFL9_from_ta.length() == 0) {
				System.out.println("DFL9_from_ta is null , continue");
				continue;
			}
			String cxbl = DFL9_from_ta;
			String[] from_tas = DFL9_from_ta.split("_");
			if(from_tas.length >= 3) {
				cxbl = from_tas[2];
			}else {
				System.out.println("DFL9_from_ta 不符合规则 , continue");
				continue;
			}
			System.out.println(taline + " -- > cxbl := " + cxbl);
			if(cxbl.equals(firstCXBL)) {
				System.out.println("相同的式样号");
				continue;
			}
			gatherTheReportData(ebomLine, area, firstCXBL, dfl9_part_no, cxbl, remark, taline, functionLine, true);
		}
	}catch(Exception e) {
		e.printStackTrace();
	}
}
private void gatherTheReportData(TCComponentBOMLine ebomLine, String area, String firstCXBL, String dfl9_part_no, String cxbl, String remark,
		TCComponentBOMLine ta , TCComponentBOMLine function, boolean isFA) {
	try {
		int k = 0;
		AIFComponentContext[] sols = ta.getChildren();//目前是在跟当前查找的同级里进行查找比较差异
		int cntSols = sols.length;
		List<TCComponentBOMLine> listSols = new ArrayList<TCComponentBOMLine>();
		for(k = 0; k < cntSols; k ++) {
			TCComponentBOMLine sol = (TCComponentBOMLine)sols[k].getComponent();
			
//			if(!listBBOMID.contains(id)) {
//				System.out.println(sol + " --> 在BBOM中不存在，不会输出到报表中" );
//				continue;
//			}
			if(sol.getItem().getType().equals("DFL9SolItmPart")) {
				listSols.add(sol);
			}
		}
		TCComponentBOMLine[] solLines = listSols.toArray(new TCComponentBOMLine[0]);
		MFCUtility.loadProperties(session, solLines, new String[] {bl_DFL9SolItmPartRevision_dfl9_part_no, "B8_MFGRemark", "bl_sequence_no", "bl_usage_address"});
		cntSols = solLines.length;
		
		
		for(k = 0; k < cntSols; k ++) {
			String part_no = solLines[k].getPropertyDisplayableValue(bl_DFL9SolItmPartRevision_dfl9_part_no);
			System.out.println("part_no := " + part_no);
			String id = solLines[k].getPropertyDisplayableValue("bl_item_item_id");
			if(part_no.equals(dfl9_part_no)) {
//				if(isFA) {
//					gatherTheReportData(ebomLine, area, firstCXBL, dfl9_part_no, cxbl, remark, 
//							solLines[k] , function, isFA) ;
//				}
				continue;//mifc 20200317
			}
//			if(part_no.equals(dfl9_part_no)) {//排除相同编号的，只保留前8位相同的
//				continue;
//			}
			if(part_no.matches(regex)) {
				System.out.println("part_no 是标准件，不查找 " );
				continue;
			}
			if(!solLines[k].getPropertyDisplayableValue("bl_usage_address").startsWith("MU")) {
				if(isFA) {
					gatherTheReportData(ebomLine, area, firstCXBL, dfl9_part_no, cxbl, remark, 
							solLines[k] , function, isFA) ;
				}
				continue;
			}
			String checkno = "";
			if(part_no.length() > 8) {
				checkno = part_no.substring(0,8);
			}
			//System.out.println("checkno := " + checkno);
			if(!StringUtil.isEmpty(checkno) && dfl9_part_no.startsWith(checkno)) {
				if(!listBBOMID.contains(id)) {
					System.out.println(solLines[k] + " --> 在BBOM中不存在，不会输出到报表中" );
					if(isFA) {
						gatherTheReportData(ebomLine, area, firstCXBL, dfl9_part_no, cxbl, remark, 
								solLines[k] , function, isFA) ;
					}
					continue;
				}
				System.out.println("dfl9_part_no get a diff no := " + part_no);
				String seqno = solLines[k].getPropertyDisplayableValue("bl_sequence_no");
				String[] rowdata = new String[6];
				rowdata[2] = cxbl;//DFL9_from_ta;
				rowdata[3] = part_no;
				rowdata[4] = solLines[k].getPropertyDisplayableValue("B8_MFGRemark");
				rowdata[5] = "";
				rowdata[0] = "";
				rowdata[1] = "";
				//rowdata[6] = seqno;
				rowdata[5] = MFCUtility.getVariantConditions(solLines[k]);
				if(this.hmFuncInfo.containsKey(function)) {
					String[] info = this.hmFuncInfo.get(function);
					rowdata[0] = info[0];
					rowdata[1] = info[1];
//					if(!this.lstArea.contains(area)) {
//						this.lstArea.add(area);
//					}
//					if(this.hmAreaFuncs.containsKey(area)) {
//						List<String> lstFu = this.hmAreaFuncs.get(area);
//						if(!lstFu.contains(info[0])) {
//							lstFu.add(info[0]);
//							this.hmAreaFuncs.put(area, lstFu);
//						}
//					}else {
//						List<String> lstFu = new ArrayList<String>();
//						lstFu.add(info[0]);
//						this.hmAreaFuncs.put(area, lstFu);
//					}
				}else {
					String[] info = new String[2];
					info[0] = function.getPropertyDisplayableValue("bl_rev_object_name");
					info[1] = function.getPropertyDisplayableValue("bl_rev_object_desc");
					rowdata[0] = info[0];
					rowdata[1] = info[1];
					hmFuncInfo.put(function, info);
//					if(!this.lstArea.contains(area)) {
//						this.lstArea.add(area);
//					}
//					if(this.hmAreaFuncs.containsKey(area)) {
//						List<String> lstFu = this.hmAreaFuncs.get(area);
//						if(!lstFu.contains(info[0])) {
//							lstFu.add(info[0]);
//							this.hmAreaFuncs.put(area, lstFu);
//						}
//					}else {
//						List<String> lstFu = new ArrayList<String>();
//						lstFu.add(info[0]);
//						this.hmAreaFuncs.put(area, lstFu);
//					}
				}
//				if(hmAFPartNums.containsKey(area + KEYINFO + rowdata[0])) {
//					List<String> lstPartnos = hmAFPartNums.get(area + KEYINFO + rowdata[0]);
//					if(!lstPartnos.contains(rowdata[3] + KEYINFO + seqno)) {
//						lstPartnos.add(rowdata[3] + KEYINFO + seqno);
//						hmAFPartNums.put(area + KEYINFO + rowdata[0], lstPartnos);
//					}
//				}else {
//					List<String> lstPartnos = new ArrayList<String>();
//					lstPartnos.add(rowdata[3] + KEYINFO + seqno);
//					hmAFPartNums.put(area + KEYINFO + rowdata[0], lstPartnos);
//				}
				StringBuffer sbKey = new StringBuffer();
				sbKey.append(area).append(KEYINFO).append(rowdata[0]).append(KEYINFO);
				sbKey.append(rowdata[2]).append(KEYINFO).append(part_no).append(KEYINFO);
				sbKey.append(rowdata[4]).append(KEYINFO).append(rowdata[5]).append(KEYINFO);
				if(!this.listReportKey.contains(sbKey.toString())) {
					if(!this.lstArea.contains(area)) {
						this.lstArea.add(area);
					}
					if(this.hmAreaFuncs.containsKey(area)) {
						List<String> lstFu = this.hmAreaFuncs.get(area);
						if(!lstFu.contains(rowdata[0])) {
							lstFu.add(rowdata[0]);
							this.hmAreaFuncs.put(area, lstFu);
						}
					}else {
						List<String> lstFu = new ArrayList<String>();
						lstFu.add(rowdata[0]);
						this.hmAreaFuncs.put(area, lstFu);
					}
					if(hmAFPartNums.containsKey(area + KEYINFO + rowdata[0])) {
						List<String> lstPartnos = hmAFPartNums.get(area + KEYINFO + rowdata[0]);
						if(!lstPartnos.contains(rowdata[3] + KEYINFO + seqno)) {
							lstPartnos.add(rowdata[3] + KEYINFO + seqno);
							hmAFPartNums.put(area + KEYINFO + rowdata[0], lstPartnos);
						}
					}else {
						List<String> lstPartnos = new ArrayList<String>();
						lstPartnos.add(rowdata[3] + KEYINFO + seqno);
						hmAFPartNums.put(area + KEYINFO + rowdata[0], lstPartnos);
					}
					listReportKey.add(sbKey.toString());
					String key = area + KEYINFO + rowdata[0] + KEYINFO + rowdata[3] + KEYINFO + seqno;
					System.out.println("key := " + key);
					this.hmAFParts.put(key, rowdata);
				}
				if(isFA) {
					gatherTheReportData(ebomLine, area, firstCXBL, dfl9_part_no, cxbl, remark, 
							solLines[k] , function, isFA) ;
				}
			}
		}
		
	}catch(Exception e) {
		e.printStackTrace();
	}
}
//private void getTheCheckPart(String area, String dfl9_part_no, String DFL9_from_ta, String remark, TCComponentBOMLine funcLine , TCComponentBOMLine solLine) {
//	try {
//			String[] rowdata = new String[6];
//			rowdata[2] = DFL9_from_ta;
//			rowdata[3] = dfl9_part_no;
//			rowdata[4] = remark;//solLine.getPropertyDisplayableValue("B8_MFGRemark");
//			rowdata[5] = "";
//			rowdata[0] = "";
//			rowdata[1] = "";
//			String seqno = solLine.getPropertyDisplayableValue("bl_sequence_no");
//			Map<String, List<String>> mapVeh = MFCUtility.getVariantCondition(solLine);
//			List<String> lstVehs = mapVeh.get("veh");
//			if(lstVehs != null) {
//				System.out.println("lstVehs.size ;= " + lstVehs.size());
//				StringBuffer sbVec = new StringBuffer();
//				for(int m = 0; m < lstVehs.size() ; m ++	) {
//					if(sbVec.toString().length() > 0) {
//						sbVec.append(",");
//					}
//					sbVec.append(lstVehs.get(m));
////					if(!this.lstVechil.contains(lstVehs.get(m))) {
////						this.lstVechil.add(lstVehs.get(m));
////					}
//				}
//				rowdata[5] = sbVec.toString();
//			}
//			if(this.hmFuncInfo.containsKey(funcLine)) {
//				String[] info = this.hmFuncInfo.get(funcLine);
//				rowdata[0] = info[0];
//				rowdata[1] = info[1];
//				if(!this.lstArea.contains(area)) {
//					this.lstArea.add(area);
//				}
//				if(this.hmAreaFuncs.containsKey(area)) {
//					List<String> lstFu = this.hmAreaFuncs.get(area);
//					if(!lstFu.contains(info[0])) {
//						lstFu.add(info[0]);
//						this.hmAreaFuncs.put(area, lstFu);
//					}
//				}else {
//					List<String> lstFu = new ArrayList<String>();
//					lstFu.add(info[0]);
//					this.hmAreaFuncs.put(area, lstFu);
//				}
//			}else {
//				String[] info = new String[2];
//				info[0] = funcLine.getPropertyDisplayableValue("bl_rev_object_name");
//				info[1] = funcLine.getPropertyDisplayableValue("bl_rev_object_desc");
//				rowdata[0] = info[0];
//				rowdata[1] = info[1];
//				hmFuncInfo.put(funcLine, info);
//				if(!this.lstArea.contains(area)) {
//					this.lstArea.add(area);
//				}
//				if(this.hmAreaFuncs.containsKey(area)) {
//					List<String> lstFu = this.hmAreaFuncs.get(area);
//					if(!lstFu.contains(info[0])) {
//						lstFu.add(info[0]);
//						this.hmAreaFuncs.put(area, lstFu);
//					}
//				}else {
//					List<String> lstFu = new ArrayList<String>();
//					lstFu.add(info[0]);
//					this.hmAreaFuncs.put(area, lstFu);
//				}
//			}
//			if(hmAFPartNums.containsKey(area + KEYINFO + rowdata[0])) {
//				List<String> lstPartnos = hmAFPartNums.get(area + KEYINFO + rowdata[0]);
//				if(!lstPartnos.contains(rowdata[3]+ KEYINFO + seqno)) {
//					lstPartnos.add(rowdata[3]+ KEYINFO + seqno);
//				}
//				hmAFPartNums.put(area + KEYINFO + rowdata[0], lstPartnos);
//			}else {
//				List<String> lstPartnos = new ArrayList<String>();
//				lstPartnos.add(rowdata[3]+ KEYINFO + seqno);
//				hmAFPartNums.put(area + KEYINFO + rowdata[0], lstPartnos);
//			}
//			String key = area + KEYINFO + rowdata[0] + KEYINFO + rowdata[3]+ KEYINFO + seqno;
//			System.out.println("key := " + key);
//			this.hmAFParts.put(key, rowdata);
//	}catch(Exception e) {
//		e.printStackTrace();
//	}
//}
//private int getFirstFA(TCComponentBOMLine[] eboms) {
//	int index = 0;
//	int i = 0;
//	int count = eboms.length ;
//	for(i = 0; i < count; i ++) {
//		if(this.hmLineFALine.get(eboms[i]) == eboms[i]) {
//			index = i;
//			break;
//		}
//	}
//	return index;
//}
//private boolean isASameFA(TCComponentBOMLine[] eboms) {
//	boolean same = true;
//	try {
//		String firstFA = eboms[0].getPropertyDisplayableValue("DFL9_from_ta");
//		int i = 1;
//		int count = eboms.length ;
//		for(i = 1; i < count; i ++) {
//			if(!eboms[i].getPropertyDisplayableValue("DFL9_from_ta").equals(firstFA)) {
//				return false;
//			}
//		}
//	}catch(Exception e) {
//		e.printStackTrace();
//	}
//	return same;
//}
private void getPartition(TCComponentBOMLine movLine, TCComponentBOMLine cline) {
	try {
		TCComponentBOMLine pline = cline.parent();
		if(pline != null && pline.getItem().getType().equals("B8_BBOMPartition")) {
			this.hmMovePartition.put(movLine, pline.getItemRevision().getProperty("object_name"));
			return;
		}else {
			getPartition(movLine, pline);
		}
	}catch(Exception e) {
		
	}
}

//	private void getFuncs(TCComponentBOMLine TALine, TCComponentBOMLine cline
//			, List<TCComponentBOMLine> lstFA, List<TCComponentBOMLine> lstFuncs) {
//		try {
//			TCComponentBOMLine pline = cline.parent();
//			if(pline != null && pline.getItem().getType().equals("DFL9Function")) {
//				this.hmLineFuncs.put(TALine, pline);
//				this.hmLineFALine.put(TALine, cline);
//				if(!lstFA.contains(cline)) {
//					lstFA.add(cline);
//				}
//				if(!lstFuncs.contains(pline)) {
//					lstFuncs.add(pline);
//				}
//				return;
//			}else {
//				getFuncs(TALine, pline, lstFA, lstFuncs);
//			}
//		}catch(Exception e) {
//			
//		}
//	}
	private void getFuncs(TCComponentBOMLine cline) {
		try {
			TCComponentBOMLine pline = cline.parent();
			if(pline != null && pline.getItem().getType().equals("DFL9Function")) {
				this.functionLine = pline;
				this.taLine = cline;
				return;
			}else {
				getFuncs(pline);
			}
		}catch(Exception e) {
			
		}
	}
}
