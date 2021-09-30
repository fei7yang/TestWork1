package com.dfl.report.dealparameter;

import java.awt.Frame;
import java.io.InputStream;
import java.math.BigDecimal;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import javax.swing.JTable;

import org.apache.log4j.Logger;
import org.eclipse.core.commands.AbstractHandler;
import org.eclipse.core.commands.ExecutionEvent;
import org.eclipse.core.commands.ExecutionException;

import com.dfl.report.ExcelReader.CurrentandVoltage;
import com.dfl.report.ExcelReader.RecommendedPressure;
import com.dfl.report.ExcelReader.SFSequenceWeldingConditionList;
import com.dfl.report.ExcelReader.SequenceComparisonTable;
import com.dfl.report.ExcelReader.SequenceWeldingConditionList;
import com.dfl.report.ExcelReader.WeldPointBoardInformation;
import com.dfl.report.ExcelReader.baseinfoExcelReader;
import com.dfl.report.util.ReportViwePanel;
import com.dfl.report.util.Util;
import com.teamcenter.rac.aif.AbstractAIFUIApplication;
import com.teamcenter.rac.aif.kernel.InterfaceAIFComponent;
import com.teamcenter.rac.aifrcp.AIFUtility;
import com.teamcenter.rac.commands.print.PropertyTablePrintDialog;
import com.teamcenter.rac.kernel.TCComponent;
import com.teamcenter.rac.kernel.TCComponentBOMLine;
import com.teamcenter.rac.kernel.TCComponentBOMWindow;
import com.teamcenter.rac.kernel.TCComponentBOMWindowType;
import com.teamcenter.rac.kernel.TCComponentItemRevision;
import com.teamcenter.rac.kernel.TCException;
import com.teamcenter.rac.kernel.TCPreferenceService;
import com.teamcenter.rac.kernel.TCSession;
import com.teamcenter.rac.util.MessageBox;

public class DealParameterHandler extends AbstractHandler {

	private AbstractAIFUIApplication application;
	private TCSession session;
	private Logger logger = Logger.getLogger(DealParameterHandler.class);
	private ArrayList discrete = new ArrayList();
	private static List<SequenceWeldingConditionList> swc = new ArrayList<SequenceWeldingConditionList>();// 24序列焊接条件设定表
																											// 序列号
	private static List<CurrentandVoltage> cv = new ArrayList<CurrentandVoltage>();// 24序列焊接条件设定表 电流电压
	private static List<SFSequenceWeldingConditionList> SFswc = new ArrayList<SFSequenceWeldingConditionList>();// 255序列焊接条件设定表
	private static List<RecommendedPressure> rp = new ArrayList<RecommendedPressure>();// 推荐加压力
	private static List<SequenceComparisonTable> sct = new ArrayList<SequenceComparisonTable>();// 序列对照表
	private Map<String, List<String>> MaterialMap;

	@Override
	public Object execute(ExecutionEvent arg0) throws ExecutionException {
		// TODO Auto-generated method stub
		System.out.println("----------------DealParameterHandler------------------");
		logger.info("----------------DealParameterHandler------------------");

		Thread thread = new Thread() {
			public void run() {
				try {
					execute();
				} catch (TCException e) {
					// TODO Auto-generated catch block
					e.printStackTrace();
				}
			}
		};
		thread.start();
		return null;
	}

	protected void execute() throws TCException {
		// TODO Auto-generated method stub
		application = AIFUtility.getCurrentApplication();
		session = (TCSession) application.getSession();

		InterfaceAIFComponent[] target = application.getTargetComponents();
		// 获取选择的对象
		InterfaceAIFComponent aifComponent = application.getTargetComponent();

		if (aifComponent == null) {
			MessageBox.post("未选择任何对象。", "提示信息", MessageBox.ERROR);
			logger.error("未选择任何对象。");
			return;
		}

		TCComponentBOMLine bomLine = null;

		// 判断选择的对象类型
		if (aifComponent instanceof TCComponentBOMLine) {
			bomLine = (TCComponentBOMLine) aifComponent;
		}

		if (bomLine == null) {
			MessageBox.post("请选择BOP中对象。", "提示信息", MessageBox.ERROR);
			logger.error("请选择BOP中对象。");
			return;
		}
		TCComponentBOMLine topbomLine = bomLine.window().getTopBOMLine();

		if (!topbomLine.getItemRevision().isTypeOf("B8_BIWPlantBOPRevision")) {
			MessageBox.post("请选择BOP中对象。", "提示信息", MessageBox.ERROR);
			logger.error("请选择BOP中对象。");
			return;
		}

		String type = bomLine.getItemRevision().getType();
		System.out.println("type:" + type);

		dealParameter(bomLine);
	}

	// 计算焊点上的参数属性，并赋值
	private void dealParameter(TCComponentBOMLine bomLine) throws TCException {
		// TODO Auto-generated method stub
		// 获取首选项定义的Note属性
		TCPreferenceService ts = session.getPreferenceService();
		if (!ts.isDefinitionExistForPreference("B8_Calculation_Parameter_Name")) {
			MessageBox.post("错误：首选项B8_Calculation_Parameter_Name未定义，请联系系统管理员！", "提示信息", MessageBox.ERROR);
			logger.error("错误：首选项B8_Calculation_Parameter_Name未定义");
			return;
		}
		// 获取材料对照表
		MaterialMap = baseinfoExcelReader.getMaterialComparisonTable(application, "DFL_MaterialMapping");
		if (MaterialMap == null || MaterialMap.size() < 1) {
			System.out.println("未找到材料对照表！");
			MessageBox.post("未配置对照表DFL_MaterialMapping，请联系系统管理员！", "提示信息", MessageBox.ERROR);
			return;
		}
		// 显示进度输出窗口
		// 界面显示进度并输出执行步骤
		ReportViwePanel viewPanel = new ReportViwePanel("焊接参数计算");
		viewPanel.setVisible(true);

		// 获取顶层
		TCComponentBOMLine topbl = bomLine.window().getTopBOMLine();

		// 获取计算参数
		Object[] obj = baseinfoExcelReader.getCalculationParameter(application, "B8_Calculation_Parameter_Name");
		if (obj != null) {
			if (obj[0] != null) {
				swc = (List<SequenceWeldingConditionList>) obj[0];
			} else {
				logger.error("未获取到24序列焊接条件设定表 序列号信息。");
				System.out.println("未获取到24序列焊接条件设定表 序列号信息。");
				viewPanel.addInfomation("未获取到24序列焊接条件设定表 序列号信息。", 100, 100);
				viewPanel.dispose();
				MessageBox.post("未找到焊接参数计算规则！", "提示信息", MessageBox.ERROR);
				return;
			}
			if (obj[1] != null) {
				cv = (List<CurrentandVoltage>) obj[1];
			} else {
				logger.error("未获取到24序列焊接条件设定表 电流电压信息。");
				System.out.println("未获取到24序列焊接条件设定表 电流电压信息。");
				viewPanel.addInfomation("未获取到24序列焊接条件设定表 电流电压信息。", 100, 100);
				viewPanel.dispose();
				MessageBox.post("未找到焊接参数计算规则！", "提示信息", MessageBox.ERROR);
				return;
			}
			if (obj[2] != null) {
				SFswc = (List<SFSequenceWeldingConditionList>) obj[2];
			} else {
				logger.error("未获取到255序列焊接条件设定表信息。");
				System.out.println("未获取到255序列焊接条件设定表信息。");
				viewPanel.addInfomation("未获取到255序列焊接条件设定表信息。", 100, 100);
				viewPanel.dispose();
				MessageBox.post("未找到焊接参数计算规则！", "提示信息", MessageBox.ERROR);
				return;
			}
			if (obj[3] != null) {
				rp = (List<RecommendedPressure>) obj[3];
			} else {
				logger.error("未获取到推荐加压力信息。");
				System.out.println("未获取到推荐加压力信息。");
				viewPanel.addInfomation("未获取到推荐加压力信息。", 100, 100);
				viewPanel.dispose();
				MessageBox.post("未找到焊接参数计算规则！", "提示信息", MessageBox.ERROR);
				return;
			}
			if (obj[4] != null) {
				sct = (List<SequenceComparisonTable>) obj[4];
			} else {
				logger.error("未获取到序列对照表信息。");
				System.out.println("未获取到序列对照表信息。");
				viewPanel.addInfomation("未获取到序列对照表信息。", 100, 100);
				viewPanel.dispose();
				MessageBox.post("未找到焊接参数计算规则！", "提示信息", MessageBox.ERROR);
				return;
			}
		} else {
			viewPanel.addInfomation("未找到焊接参数计算规则！", 100, 100);
			logger.error("未找到焊接参数计算规则！");
			viewPanel.dispose();
			MessageBox.post("未找到焊接参数计算规则！", "提示信息", MessageBox.ERROR);
			return;
		}
		viewPanel.addInfomation("获取相关参数信息完成...\n", 20, 100);

		List<WeldPointBoardInformation> weldboradlist = getBaseinfomation(topbl, "222.基本信息");
		if (weldboradlist == null || weldboradlist.size() < 1) {
			viewPanel.addInfomation("基本信息为空，请确认是否已生成基本信息！", 100, 100);
			logger.error("基本信息为空，请确认是否已生成基本信息！");
			viewPanel.dispose();
			MessageBox.post("基本信息为空，请确认是否已生成基本信息！", "提示信息", MessageBox.ERROR);
			return;
		}
		viewPanel.addInfomation("获取基本信息完成\n", 40, 100);
		discrete = getAllDiscrete(bomLine);
		if (discrete == null || discrete.size() < 1) {

			viewPanel.addInfomation("点焊工序数据为空，请确认所选对象下存在点焊工序数据！", 100, 100);
			logger.error("点焊工序数据为空，请确认所选对象下存在点焊工序数据！");
			viewPanel.dispose();
			MessageBox.post("点焊工序数据为空，请确认所选对象下存在点焊工序数据！", "提示信息", MessageBox.ERROR);
			return;
		}
		viewPanel.addInfomation("获取焊点信息完成...\n", 60, 100);

		// 循环焊点信息，计算并获取参数属性值
		Map<TCComponentBOMLine, String[]> paramap = new HashMap<TCComponentBOMLine, String[]>();
		// 输出厚薄板件，提示用户
		ArrayList bzlist = new ArrayList();
		// 用于判断焊点重复，导致重复执行
		ArrayList cflist = new ArrayList();
		// 用于记录无焊枪的工序名称
		String discretenames = "";
		// 用于记录无区分伺服和气动属性的焊枪名称
		String gunnames = "";

		for (int i = 0; i < discrete.size(); i++) {

			TCComponentBOMLine bl = (TCComponentBOMLine) discrete.get(i);
			ArrayList weld = Util.getChildrenByBOMLine(bl, "WeldPointRevision");
			ArrayList gun = Util.getChildrenByBOMLine(bl, "B8_BIWGunRevision");
			String direname = Util.getProperty(bl, "bl_rev_object_name");
			// 只有R开头的才需要判断
			String GunType = "";
			if (direname.length() > 1) {
				if (direname.substring(0, 1).equals("R")) {
					// 只有工序下有焊点才判断有没有枪
					if (weld != null && weld.size() > 0) {
						if (gun != null && gun.size() > 0) {
							TCComponentBOMLine gbl = (TCComponentBOMLine) gun.get(0);
							GunType = Util.getProperty(gbl.getItemRevision(), "b8_GunType");
							if (GunType.isEmpty()) {
								if (gunnames.isEmpty()) {
									gunnames = direname;
								} else {
									gunnames = gunnames + "," + direname;
								}
							}
						} else {
							if (weld != null && weld.size() > 0) {
								if (discretenames.isEmpty()) {
									discretenames = direname;
								} else {
									discretenames = discretenames + "," + direname;
								}
							}
						}
					}

				}
			}

			if (weld != null && weld.size() > 0) {
				for (int k = 0; k < weld.size(); k++) {
					TCComponentBOMLine wbl = (TCComponentBOMLine) weld.get(k);
					String weldno = Util.getProperty(wbl, "bl_rev_object_name");
					String[] str = new String[13];
					for (int j = 0; j < weldboradlist.size(); j++) {
						WeldPointBoardInformation wbinfo = weldboradlist.get(j);
						String wbNO = wbinfo.getWeldno();
						if (wbNO.equals(weldno)) {
							if (!cflist.contains(weldno)) {
								cflist.add(weldno);
								//非gagi，根据材质对照表，判断焊点是否需要计算参数
								String Partmaterial = wbinfo.getPartmaterial1();								
								String Partmateria2 = wbinfo.getPartmaterial2();								
								String Partmateria3 = wbinfo.getPartmaterial3();
								String gagi1 = wbinfo.getGagi1();								
								String gagi2 = wbinfo.getGagi2();								
								String gagi3 = wbinfo.getGagi3();
								boolean isCalculate = true;
								if(MaterialMap!=null)
								{
									for(Map.Entry<String, List<String>> entry: MaterialMap.entrySet())
									{
										String MaterialNo = entry.getKey();
										List<String> infolist = entry.getValue();
										if(!"GA".equals(gagi1) && !"GI".equals(gagi1))
										{
											if(Util.getIsEqueal(Partmaterial, MaterialNo))
											{
												if("否".equalsIgnoreCase(infolist.get(1)))
												{
													isCalculate = false;
													break;
												}								
											}
										}
										if(!"GA".equals(gagi2) && !"GI".equals(gagi2))
										{
											if(Util.getIsEqueal(Partmateria2, MaterialNo))
											{
												if("否".equalsIgnoreCase(infolist.get(1)))
												{
													isCalculate = false;
													break;
												}								
											}
										}
										if(!"GA".equals(gagi3) && !"GI".equals(gagi3))
										{
											if(Util.getIsEqueal(Partmateria3, MaterialNo))
											{
												if("否".equalsIgnoreCase(infolist.get(1)))
												{
													isCalculate = false;
													break;
												}								
											}
										}																							
									}
								}
								if(!isCalculate)
								{
									str[0] = "";
									str[1] = "";
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
									str[12] = "";
									paramap.put(wbl, str);
									System.out.println(weldno + "焊点存在不需要参与计算关联板件！");
									break; //该焊点存在不参与的板件
								}
								
								String boradnum = wbinfo.getLayersnum();// 板层数
								String basethickness = wbinfo.getBasethickness();// 基准板厚
								String sheetstrength1 = wbinfo.getStrength1();// 板件1强度
								String sheetstrength2 = wbinfo.getStrength2();// 板件2强度
								String sheetstrength3 = wbinfo.getStrength3();// 板件3强度
								String partthickness1 = wbinfo.getPartthickness1();// 板件1厚度
								String partthickness2 = wbinfo.getPartthickness2();// 板件2厚度
								String partthickness3 = wbinfo.getPartthickness3();// 板件3厚度
//								String gagi1 = wbinfo.getGagi1();// 板件1GA/GI材
//								String gagi2 = wbinfo.getGagi2();// 板件2GA/GI材
//								String gagi3 = wbinfo.getGagi3();// 板件3GA/GI材

								// 推荐加压力属性
								String Repressure = "";
								Repressure = getRepressure(basethickness, boradnum, sheetstrength1, sheetstrength2,
										sheetstrength3);
								str[0] = Repressure;
								// System.out.println("焊点"+weldno+"推荐加压力为："+Repressure);
								// 24序列焊接条件设定表 参数序列号
								String parameterSerialNo24 = "";
								parameterSerialNo24 = getParameterSerialNo24(basethickness, boradnum, gagi1, gagi2,
										gagi3, sheetstrength1, sheetstrength2, sheetstrength3);

								str[1] = parameterSerialNo24;
								// 255序列焊接条件设定表 参数序列号
								// 获取板厚度差
								double thicknessdifference = getThicknessDifference(partthickness1, partthickness2,
										partthickness3, boradnum);
								String parameterSerialNo255 = "";
								parameterSerialNo255 = getParameterSerialNo255(basethickness, boradnum, gagi1, gagi2,
										gagi3, sheetstrength1, sheetstrength2, sheetstrength3, thicknessdifference);
								// 电流参数序列（日产） 需要区分气动和伺服 ，逻辑待定
								if (GunType.equals("伺服")) {
									// 只有RSW伺服计算推荐序列（对应）
									str[1] = parameterSerialNo255;
									System.out.println("焊点" + weldno + " 参数序列(日产)：" + parameterSerialNo255);
									// 255序列对照表
									String SequenceComparison = "";
									SequenceComparison = getSequenceComparison(parameterSerialNo255);
									System.out.println("焊点" + weldno + " 参数序列(对应)：" + parameterSerialNo255);
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
									// 只有PSW和RSW启动需要计算电流值
									// 24序列焊接条件设定表 推荐 电流值
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

								String partmaterial1 = wbinfo.getPartmaterial1();
								String partmaterial2 = wbinfo.getPartmaterial2();
								String partmaterial3 = wbinfo.getPartmaterial3();

								String partNos = "";
								// 判断是否为厚薄板
								boolean flag1 = getJudgingThickSheet(partmaterial1);
								if (flag1) {
									String partNo1 = wbinfo.getPartNo1();
									if (!bzlist.contains(partNo1)) {
										if (partNos.isEmpty()) {
											partNos = partNo1;
										} else {
											partNos = partNos + "," + partNo1;
										}
									}
								}
								boolean flag2 = getJudgingThickSheet(partmaterial2);
								if (flag2) {
									String partNo2 = wbinfo.getPartNo2();
									if (!bzlist.contains(partNo2)) {
										if (partNos.isEmpty()) {
											partNos = partNo2;
										} else {
											partNos = partNos + "," + partNo2;
										}
									}
								}
								boolean flag3 = getJudgingThickSheet(partmaterial3);
								if (flag3) {
									String partNo3 = wbinfo.getPartNo3();
									if (!bzlist.contains(partNo3)) {
										if (partNos.isEmpty()) {
											partNos = partNo3;
										} else {
											partNos = partNos + "," + partNo3;
										}
									}
								}
								if (flag1 || flag2 || flag3) {
									String[] strValues = new String[2];
									strValues[0] = weldno;
									strValues[1] = partNos;
									bzlist.add(strValues);
								} else {
									paramap.put(wbl, str);
								}
								break;
							}
						}
					}
				}
			}

		}
		if (!discretenames.isEmpty()) {
			viewPanel.dispose();
			MessageBox.post("点焊工序：" + discretenames + "下没有焊枪，无法计算焊接参数！", "提示信息", MessageBox.ERROR);
			return;
		}
		if (!gunnames.isEmpty()) {
			viewPanel.dispose();
			MessageBox.post("点焊工序：" + gunnames + "下枪的区分伺服/气动属性为空，无法计算焊接参数！", "提示信息", MessageBox.ERROR);
			return;
		}

		viewPanel.addInfomation("焊接参数计算中...\n", 80, 100);

		// 根据属性参数值，写属性值
		String[] properties = { "bl_WeldPointRevision_b8_RecomWeldForce", "bl_WeldPointRevision_b8_CurrentSerie_Nissan",
				"bl_WeldPointRevision_b8_RiseTime", "bl_WeldPointRevision_b8_CurrentTime1",
				"bl_WeldPointRevision_b8_Current1", "bl_WeldPointRevision_b8_Cool1",
				"bl_WeldPointRevision_b8_CurrentTime2", "bl_WeldPointRevision_b8_Current2",
				"bl_WeldPointRevision_b8_Cool2", "bl_WeldPointRevision_b8_CurrentTime3",
				"bl_WeldPointRevision_b8_Current3", "bl_WeldPointRevision_b8_KeepTime",
				"bl_WeldPointRevision_b8_CurrentSerie_DFL" };

		{
			// 开启旁路
			Util.callByPass(session, true);
		}
		// 写属性值
		Util.setAllCompsProperty(session, paramap, properties);
		{
			// 关闭旁路
			Util.callByPass(session, false);
		}
		viewPanel.addInfomation("焊接参数属性值写入完成。", 100, 100);

		if (bzlist != null && bzlist.size() > 0) {
			viewPanel.dispose();
			printUnAssigned(bzlist);
		}
	}

	public void printUnAssigned(ArrayList list) {
		if (list != null && !list.isEmpty()) {

			int count = list.size();
			String[] serializable;
			Object[][] arrobject = new Object[count][2];
			int n = 0;
			while (n < count) {
				serializable = (String[]) list.get(n);
				arrobject[n][0] = serializable[0];
				arrobject[n][1] = serializable[1];
				++n;
			}
			Object[] arrobject2 = new String[] { "焊点ID", "厚薄板件" };
			Frame frame = AIFUtility.getActiveDesktop().getFrame();
			PropertyTablePrintDialog propertyTablePrintDialog = new PropertyTablePrintDialog(frame,
					new JTable(arrobject, arrobject2), "焊点清单");
			propertyTablePrintDialog.setTitle("未计算焊接参数焊点清单");
			propertyTablePrintDialog.setVisible(true);
			propertyTablePrintDialog.setAlwaysOnTop(true);
		}
	}

	// 判断是否为厚薄板
	private boolean getJudgingThickSheet(String partmaterial1) {
		// TODO Auto-generated method stub
		boolean flag = false;
		int count1 = 0;
		int count2 = 0;
		String str = "";
		if (partmaterial1 != null) {
			str = partmaterial1;
		}
		count1 = (str.length() - str.replace("SP", "").length()) / "SP".length();
		count2 = (str.length() - str.replace("RP", "").length()) / "RP".length();

		if (count1 + count2 > 1) {
			flag = true;
		}
		return flag;
	}

	// 255序列对照表
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

	// 255序列焊接条件设定表 参数序列号
	public static String getParameterSerialNo255(String basethickness, String boradnum, String gagi1, String gagi2,
			String gagi3, String sheetstrength1, String sheetstrength2, String sheetstrength3,
			double thicknessdifference) {
		// TODO Auto-generated method stub
		String parameterSerialNo255 = "";
		int lnum = 0; // 裸板数量
		int ganum = 0; // GA材数量
		int high = 0;// 高强材数量
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

	// 获取板厚度差
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

	// 24序列焊接条件设定表 推荐 电流值
	public static String[] getRecommendedvalue(String parameterSerialNo24) {
		// TODO Auto-generated method stub
		String[] recommendedvalue = new String[10];
		for (int i = 0; i < cv.size(); i++) {
			CurrentandVoltage cvotage = cv.get(i);
			String serialNo = cvotage.getSequenceNo();
			if (serialNo != null && serialNo.equals(parameterSerialNo24)) {
				recommendedvalue[0] = cvotage.getBvalue();// 上升时间
				recommendedvalue[1] = cvotage.getCvalue();// 第一 通电时间
				recommendedvalue[2] = cvotage.getEvalue();// 第一 通电电流
				recommendedvalue[3] = cvotage.getFvalue();// 冷却时间一
				recommendedvalue[4] = cvotage.getGvalue();// 第二通电时间
				recommendedvalue[5] = cvotage.getIvalue();// 第二通电电流
				recommendedvalue[6] = cvotage.getJvalue();// 冷却时间二
				recommendedvalue[7] = cvotage.getKvalue();// 第三 通电时间
				recommendedvalue[8] = cvotage.getMvalue();// 第三 通电电流
				recommendedvalue[9] = cvotage.getNvalue();// 保持
				break;
			}
		}
		return recommendedvalue;
	}

	// 24序列焊接条件设定表 参数序列号
	public static String getParameterSerialNo24(String basethickness, String boradnum, String gagi1, String gagi2,
			String gagi3, String sheetstrength1, String sheetstrength2, String sheetstrength3) {
		// TODO Auto-generated method stub
		String parameterSerialNo24 = "";
		int lnum = 0; // 裸板数量
		int ginum = 0; // GI材数量
		int ganum = 0; // GA材数量
		int high = 0;// 高强材数量

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
					// 当板板组中既有GA材又有GI材时，将GA材当做GI材来考虑。

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

	// 获取加压力
	public static String getRepressure(String basethickness, String boradnum, String sheetstrength1,
			String sheetstrength2, String sheetstrength3) {
		// TODO Auto-generated method stub
		String repressure = "";
		String distinguish = ""; // 区分
		int num1 = 0;// 440以下数量
		int num2 = 0;// 440
		int num3 = 0;// 590Mpa780Mpa980Mpa
		// 如果存在1180强度板材，不计算参数列，默认为空
		int shstrength1 = getInteger(sheetstrength1);
		int shstrength2 = getInteger(sheetstrength2);
		int shstrength3 = getInteger(sheetstrength3);
		if (shstrength1 == 1180 || shstrength2 == 1180 || shstrength3 == 1180) {
			repressure = "";
			return repressure;
		}
		// 如果存在1350强度板材，区分为I
		if (shstrength1 == 1350 || shstrength2 == 1350 || shstrength3 == 1350) {
			distinguish = "Ⅰ";
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

			// 先根据两层板规则，获取分区,再根据3层板
			if (boradnum.equals("2")) {
				if (num1 == 3) {
					distinguish = "Ⅰ";
				}
				if (num3 == 0 && num2 == 1) {
					distinguish = "Ⅱ";
				}
				if (num2 == 0 && num3 == 1) {
					distinguish = "Ⅲ";
				}
				if (num2 == 2) {
					distinguish = "Ⅲ";
				}
				if (num2 == 1 && num3 == 1) {
					distinguish = "Ⅳ";
				}
				if (num3 == 2) {
					distinguish = "Ⅴ";
				}
			} else if (boradnum.equals("3")) {
				if (num1 == 3) {
					distinguish = "Ⅰ";
				}
				if (num1 == 2 && num2 == 1) {
					distinguish = "Ⅱ";
				}
				if (num1 == 2 && num3 == 1) {
					distinguish = "Ⅲ";
				}
				if (num1 == 1 && num2 == 2) {
					distinguish = "Ⅲ";
				}
				if (num1 == 1 && num2 == 1 && num3 == 1) {
					distinguish = "Ⅳ";
				}
				if (num1 == 1 && num3 == 2) {
					distinguish = "Ⅴ";
				}
				if (num2 == 3) {
					distinguish = "Ⅲ";
				}
				if (num2 == 2 && num3 == 1) {
					distinguish = "Ⅴ";
				}
				if (num2 == 1 && num3 == 2) {
					distinguish = "Ⅴ";
				}
				if (num3 == 3) {
					distinguish = "Ⅴ";
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
					if (distinguish.equals("Ⅰ")) {
						repressure = repre.getBvalue();
					}
					if (distinguish.equals("Ⅱ")) {
						repressure = repre.getCvalue();
					}
					if (distinguish.equals("Ⅲ")) {
						repressure = repre.getDvalue();
					}
					if (distinguish.equals("Ⅳ")) {
						repressure = repre.getEvalue();
					}
					if (distinguish.equals("Ⅴ")) {
						repressure = repre.getFvalue();
					}
					break;
				}
			}

		}

		return repressure;
	}

	// 获取BOP下基本信息中焊点板组信息
	public static List<WeldPointBoardInformation> getBaseinfomation(TCComponentBOMLine topbl, String procName) {
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

	// 获取BOP中的所有焊点
	public ArrayList<TCComponentBOMLine> getAllDiscrete(TCComponentBOMLine topbl) {
		// TODO Auto-generated method stub

		ArrayList<TCComponentBOMLine> welds = new ArrayList<TCComponentBOMLine>();
		System.out.println("用于测试无响应问题");

		ArrayList qclist = new ArrayList();
		// 根据BOP查询所有的点焊工序
		String typename = Util.getObjectDisplayName(session, "B8_BIWDiscreteOP");
		String[] propertys = new String[] { "bl_item_object_type", "bl_item_object_type" };
		String[] values = new String[] { typename, typename };

		ArrayList partList = Util.searchBOMLine(topbl, "OR", propertys, "==", values);
		System.out.println("点焊工序集合：" + partList.toString());

		return partList;
	}

	/*
	 * 取最小值
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
	 * 取最大值
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
	 * 字符转换成整数
	 */
	public static int getInteger(String str) {
		int num = -1;
		if (Util.isNumber(str)) {
			num = (int) Double.parseDouble(str);
		}
		return num;
	}

}
