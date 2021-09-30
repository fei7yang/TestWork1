package com.dfl.report.handlers;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.InputStream;
import java.text.DecimalFormat;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.HashMap;
import java.util.List;
import java.util.logging.Logger;

import javax.swing.SwingUtilities;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.dfl.report.ExcelReader.CoverInfomation;
import com.dfl.report.ExcelReader.baseinfoExcelReader;
import com.dfl.report.util.FileUtil;
import com.dfl.report.util.NewOutputDataToExcel;
import com.dfl.report.util.ReportViwePanel;
import com.dfl.report.util.Util;
import com.teamcenter.rac.aif.AbstractAIFUIApplication;
import com.teamcenter.rac.aif.kernel.AIFComponentContext;
import com.teamcenter.rac.aif.kernel.InterfaceAIFComponent;
import com.teamcenter.rac.kernel.TCComponent;
import com.teamcenter.rac.kernel.TCComponentBOMLine;
import com.teamcenter.rac.kernel.TCComponentDataset;
import com.teamcenter.rac.kernel.TCComponentDatasetType;
import com.teamcenter.rac.kernel.TCComponentItemRevision;
import com.teamcenter.rac.kernel.TCComponentTcFile;
import com.teamcenter.rac.kernel.TCException;
import com.teamcenter.rac.kernel.TCPreferenceService;
import com.teamcenter.rac.kernel.TCSession;
import com.teamcenter.rac.util.MessageBox;

public class StraightforwardlistOp {

	private AbstractAIFUIApplication app;
	private ReportViwePanel viewPanel;
	// 大小件规则
	private HashMap<String, String> rule = new HashMap<String, String>();
	SimpleDateFormat df = new SimpleDateFormat("yyyyMMdd  HH");// 设置日期格式
	private TCComponent folder;
	private static Logger logger = Logger.getLogger(baseinfoExcelReader.class.getName()); // 日志打印类
	private InterfaceAIFComponent[] ifc;
	private TCSession session;
	private InputStream inputStream;

	public StraightforwardlistOp(AbstractAIFUIApplication app, TCComponent savefolder, InterfaceAIFComponent[] ifc,
			TCSession session, InputStream inputStream) {
		// TODO Auto-generated constructor stub
		this.app = app;
		this.folder = savefolder;
		this.ifc = ifc;
		this.session = session;
		this.inputStream = inputStream;
		initUI();
	}

	// 变量
	int hsnum = 0;// 焊点总数
	int rswnum = 0;// RSW总数

	private void initUI() {
		// TODO Auto-generated method stub
		try {

			// InterfaceAIFComponent[] aifc = app.getTargetComponents();

			TCComponentBOMLine aifbl = (TCComponentBOMLine) ifc[0];
			TCComponentBOMLine topbl = aifbl.window().getTopBOMLine();

			String familiycode = Util.getProperty(topbl.getItemRevision(), "project_ids");
			String vecile = Util.getDFLProjectIdVehicle(familiycode);
			if(vecile==null || vecile.isEmpty()) {
				vecile = familiycode;
			}

//			// 根据顶层BOP查询所有的焊装产线
//			String[] propertys = new String[] { "bl_item_object_type", "bl_item_object_type" };
//			String[] values = new String[] { "焊装产线工艺", "BIW Process Line" };
//
//			// 根据焊装产线查询所有的点焊工序
//			String[] propertys2 = new String[] { "bl_item_object_type", "bl_item_object_type" };
//			String[] values2 = new String[] { "焊装工位工艺", "BIW Process Station" };

			viewPanel = new ReportViwePanel("生成报表");
			viewPanel.setVisible(true);

			// viewPanel.addInfomation("开始输出报表...\n", 5, 100);
			viewPanel.addInfomation("开始输出报表......\n", 10, 100);

			
			// 先获取所有的虚层产线
			List CList = Util.getChildrenByParent(ifc);

			if (CList == null) {
				viewPanel.dispose();
				MessageBox.post("错误：当前车型的PlantBOP下，没有点焊工序数据！", "温馨提示", MessageBox.INFORMATION);				
				return;
			}
			// 把METAL产线放到最后输出
			ArrayList SortList = getOrderList(CList);

			// 再根据所有虚层产线获取实际产线
			// ArrayList partList = getFactLineByParent(SortList);

			if (SortList == null) {
				viewPanel.dispose();
				MessageBox.post("错误：当前车型的PlantBOP下，没有点焊工序数据！", "温馨提示", MessageBox.INFORMATION);
				return;
			}

			viewPanel.addInfomation("", 20, 100);

			XSSFWorkbook book = NewOutputDataToExcel.creatXSSFWorkbook(inputStream);

			// 汇总行的行数标记
			int total_rownum = 11;// 获取模板的初始位置

			// 获取大小件规则
			rule = getSizeRule();

			viewPanel.addInfomation("开始写数据，请耐心等待...\n", 40, 100);
			for (int i = 0; i < SortList.size(); i++) {

				viewPanel.addInfomation("", 40, 100);

				TCComponentBOMLine topline = (TCComponentBOMLine) SortList.get(i);
				// 获取工位集合
				ArrayList discreteList = getStateChildrenByParent(topline);

				/*
				 * ***************************** 根据虚产线的名称去判断是否为区分机器焊还是人工焊
				 */
				int ajm = getIsAJ(topline);
				// 虚层产线名称 需要拼接线体式样输出 20191202修改				
				String plinename = topline.getProperty("bl_rev_object_name");
				long startTime = System.currentTimeMillis(); // 获取开始时间
				ArrayList list = getArrayListData(discreteList, ajm);
				long endTime = System.currentTimeMillis(); // 获取结束时间
				System.out.println("获取一个产线数据： " + (endTime - startTime) + "ms");

				list.add(plinename);

				System.out.println("产线名称：" + list.get(list.size() - 1) + "/产线下的数据行数：" + list.get(list.size() - 2));

				NewOutputDataToExcel.writeDataToSheet(book, list, hsnum, rswnum, total_rownum, viewPanel, ajm);

				if (total_rownum == 11) {
					total_rownum = total_rownum + (int) (list.get(list.size() - 2)) - 1;
				} else {
					total_rownum = total_rownum + (int) (list.get(list.size() - 2));
				}

				System.out.println("总行数：" + total_rownum);
			}
			viewPanel.addInfomation("", 60, 100);
			// 最后删除模板行并设置汇总行的公式
			NewOutputDataToExcel.dealTotalRowFormula(book, viewPanel);

			String date = df.format(new Date());
			String datasetname = vecile + "直劳计算表" + "_" + date + "时";
			String filename = Util.formatString(datasetname);

			NewOutputDataToExcel.exportFile(book, filename);

			viewPanel.addInfomation("", 80, 100);

			// String fullFileName = FileUtil.getReportFileName("直劳清单表");

			Util.saveFiles(filename, datasetname, folder, session, "AD");
			// NewOutputDataToExcel.openFile(fullFileName);

			viewPanel.addInfomation("输出报表完成，请在选择保存的文件夹下查看！\n", 100, 100);

		} catch (Exception e) {
			e.printStackTrace();
		}
	}

	private ArrayList getStateChildrenByParent(TCComponentBOMLine topline) {
		// TODO Auto-generated method stub
		ArrayList list = new ArrayList();
		try {
			AIFComponentContext[] childrens = topline.getChildren();
			for (AIFComponentContext chil : childrens) {
				TCComponentBOMLine bl = (TCComponentBOMLine) chil.getComponent();
				AIFComponentContext[] childrens2 = bl.getChildren();
				for (AIFComponentContext chil2 : childrens2) {
					TCComponentBOMLine dbl = (TCComponentBOMLine) chil2.getComponent();
					if(dbl.getItemRevision().isTypeOf("B8_BIWMEProcStatRevision")) {
						list.add(dbl);
					}				
				}
			}
			return list;

		} catch (TCException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		return list;
	}

	private ArrayList getFactLineByParent(ArrayList cList) throws TCException {
		// TODO Auto-generated method stub
		ArrayList list = new ArrayList();
		for (int i = 0; i < cList.size(); i++) {
			TCComponentBOMLine pbl = (TCComponentBOMLine) cList.get(i);
			AIFComponentContext[] childrens = pbl.getChildren();
			for (AIFComponentContext aif : childrens) {
				TCComponentBOMLine bl = (TCComponentBOMLine) aif.getComponent();
				TCComponentItemRevision rev = bl.getItemRevision();
				if (rev.isTypeOf("B8_BIWMEProcLineRevision")) {
					list.add(bl);
				}
			}
		}
		return list;
	}

	// 获取产线下的点焊工序
	private ArrayList getArrayListData(ArrayList partList, int ajm) {

		ArrayList list = new ArrayList();// 主数据集合
		try {

			String PARTS_NAME;// PARTS NAME(OPERATION NAME)列值

			int GUN;// GUN QTY列值
			int RSW;// RSW(MSW)列值
			int PSW;// PSW PTS列值
			int rownum = 0;// 汇总行需要向下移动的行数
			/*
			 * 根据焊装工位工艺是否有多个点焊工序，如果有PARTS NAME(OPERATION NAME)
			 * 列则取值为焊装工位工艺（点焊工序名称），否则只取值焊装工位工艺；针对JR打头的工位工艺，如果
			 * 有多个点焊工序，分点焊工序输出焊点信息，（）为点焊工序名称。如果是人工焊接工位（AJ打头），则不区分工序，合并输出
			 */
			for (int i = 0; i < partList.size(); i++) {
				viewPanel.addInfomation("", 40, 100);

				ArrayList ajlist = new ArrayList();// 点焊工序集合

				TCComponentBOMLine bl = (TCComponentBOMLine) partList.get(i);
				String plinename = Util.getProperty(bl.parent().getItemRevision(), "b8_LineType")
						+ bl.parent().getProperty("bl_rev_object_name");

				boolean flag = getIsDiscretes(bl); // 判断是否存在含有以R开头的机器人工序
				// 其他类不需要区分机器人工，按工位统计输出
				if (ajm == 3) {
					int LARGE_PARTS = 0;// LARGE PARTS列值
					int SUPERLARGE_PARTS = 0;// SUPERLARGE_PARTS列值
					int MID_PARTS = 0;// MID PARTS列值
					int SMALL_PARTS = 0;// SMALL PARTS列值
					int CLAMP = 0;// CLAMP& UN-CLAMP列值

					String[] str = new String[10];// 共10个数据
					PARTS_NAME = plinename + " " + bl.getProperty("bl_rev_object_name");

					ajlist = Util.getChildrenByParent(bl);

					str[0] = Integer.toString(rownum + 1);// 序号
					str[1] = PARTS_NAME;

					// 根据焊装工位查询零件
					ArrayList part = new ArrayList();// 零件集合

					getSolItmPart(bl, part);

					for (int j = 0; j < part.size(); j++) {
						TCComponentItemRevision rev = (TCComponentItemRevision) part.get(j);
						String partno = Util.getProperty(rev, "dfl9_part_no");
						if (partno.length() > 5) {
							partno = partno.substring(0, 5);
						}
						if (rule.containsKey(partno)) {
							String type = rule.get(partno);
							if (type.equals("SUPER LARGE PARTS")) {
								SUPERLARGE_PARTS++;
							}
							if (type.equals("LARGE PARTS")) {
								LARGE_PARTS++;
							}
							if (type.equals("MID PARTS")) {
								MID_PARTS++;
							}
							if (type.equals("SMALL PARTS")) {
								SMALL_PARTS++;
							}
						}
					}

					if (ajlist != null) {

						// SUPERLARGE_PARTS = 0;
						if (SUPERLARGE_PARTS != 0) {
							str[2] = Integer.toString(SUPERLARGE_PARTS);
						}
						// LARGE_PARTS = 0;
						if (LARGE_PARTS != 0) {
							str[3] = Integer.toString(LARGE_PARTS);
						}
						// MID_PARTS = 0;
						if (MID_PARTS != 0) {
							str[4] = Integer.toString(MID_PARTS);
						}
						// SMALL_PARTS = 0;
						if (SMALL_PARTS != 0) {
							str[5] = Integer.toString(SMALL_PARTS);
						}
						if (CLAMP != 0) {
							str[6] = Integer.toString(CLAMP);
						}
						// CLAMP = 0;

						// 根据焊装工位去统计枪和焊点
//						ArrayList QList = Util.searchBOMLine(bl, "OR",
//								new String[] { "bl_item_object_type", "bl_item_object_type" }, "==",
//								new String[] { "枪", "BIW Gun" });
//						ArrayList QList = getGunNums(bl);						
//						ArrayList HDList = Util.searchBOMLine(bl, "OR",
//								new String[] { "bl_item_object_type", "bl_item_object_type" }, "==",
//								new String[] { "焊点", "Weld Point" });
						Integer[] hdnums = getWeldPointNums(bl);

//						if (QList != null) {
//							GUN = QList.size();
//							str[7] = Integer.toString(GUN);
//						} else {
//							str[7] = "0";
//						}
						str[7] = getGunNums(bl).toString();
						if (hdnums != null) {
							PSW = hdnums[0];
							RSW = hdnums[1];
							if (RSW != 0) {
								str[8] = hdnums[1].toString();
							}
							if (PSW != 0) {
								str[9] = hdnums[0].toString();
							}
							hsnum = hsnum + PSW + RSW;
							rswnum = rswnum + RSW;
						}
//						if (HDList != null) {
//							PSW = HDList.size();
//							str[9] = Integer.toString(PSW);
//							hsnum = hsnum + PSW;
//						} else {
//							PSW = 0;
//							str[9] = Integer.toString(PSW);
//						}
					}
					rownum++;
					list.add(str);

				}
				// 需要区分机器人工产线，且需要注意上料工序和点焊工序如果名称一样，需要合并一行输出
				if (ajm == 1) {
					// 工位中有机器人工序
					if (flag) {
						ajlist = Util.getChildrenByBOMLine(bl, "B8_BIWDiscreteOPRevision");// 获取工位下的点焊工序
						ArrayList oplist = Util.getChildrenByBOMLine(bl, "B8_BIWOperationRevision");// 获取工位下的上料工序
						for (int j = 0; j < ajlist.size(); j++) {
							int LARGE_PARTS = 0;// LARGE PARTS列值
							int SUPERLARGE_PARTS = 0;// SUPERLARGE_PARTS列值
							int MID_PARTS = 0;// MID PARTS列值
							int SMALL_PARTS = 0;// SMALL PARTS列值
							int CLAMP = 0;// CLAMP& UN-CLAMP列值

							String[] str = new String[10];// 共10个数据
							TCComponentBOMLine chil = (TCComponentBOMLine) ajlist.get(j);
							String discretename = chil.getProperty("bl_rev_object_name");
							TCComponentBOMLine relatively = null; // 对应的上料工序
							// 如果点焊工序名称和上件工序名称一致，合并输出
							if (oplist != null) {
								for (int m = 0; m < oplist.size(); m++) {
									TCComponentBOMLine ch = (TCComponentBOMLine) oplist.get(m);
									String opname = ch.getProperty("bl_rev_object_name");
									if (discretename.equals(opname)) {
										relatively = ch;
										oplist.remove(m);
										break;
									}
								}
							}

							// 根据焊装工位查询零件
							ArrayList part = new ArrayList();// 零件集合

							if (relatively != null) {
								getSolItmPart(relatively, part);
							}

							for (int k = 0; k < part.size(); k++) {
								TCComponentItemRevision rev = (TCComponentItemRevision) part.get(k);
								String partno = Util.getProperty(rev, "dfl9_part_no");
								if (partno.length() > 5) {
									partno = partno.substring(0, 5);
								}
								if (rule.containsKey(partno)) {
									String type = rule.get(partno);
									if (type.equals("SUPER LARGE PARTS")) {
										SUPERLARGE_PARTS++;
									}
									if (type.equals("LARGE PARTS")) {
										LARGE_PARTS++;
									}
									if (type.equals("MID PARTS")) {
										MID_PARTS++;
									}
									if (type.equals("SMALL PARTS")) {
										SMALL_PARTS++;
									}
								}
							}

							PARTS_NAME = plinename + " " + bl.getProperty("bl_rev_object_name") + "(" + discretename
									+ ")";

							str[0] = Integer.toString(rownum + 1);// 序号

							str[1] = PARTS_NAME;

							// SUPERLARGE_PARTS = 0;
							if (SUPERLARGE_PARTS != 0) {
								str[2] = Integer.toString(SUPERLARGE_PARTS);
							}
							// LARGE_PARTS = 0;
							if (LARGE_PARTS != 0) {
								str[3] = Integer.toString(LARGE_PARTS);
							}
							// MID_PARTS = 0;
							if (MID_PARTS != 0) {
								str[4] = Integer.toString(MID_PARTS);
							}
							// SMALL_PARTS = 0;
							if (SMALL_PARTS != 0) {
								str[5] = Integer.toString(SMALL_PARTS);
							}
							if (CLAMP != 0) {
								str[6] = Integer.toString(CLAMP);
							}

							ArrayList QList = Util.getChildrenByBOMLine(chil, "B8_BIWGunRevision");
							ArrayList HDList = Util.getChildrenByBOMLine(chil, "WeldPointRevision");
							// 根据点焊工序的名称是否为R开头，确定机器还是人工
							if (discretename.substring(0, 1).equals("R")) {
								RSW = HDList.size();
								str[8] = Integer.toString(RSW);
								hsnum = hsnum + RSW;
								rswnum = rswnum + RSW;
							} else {
								if (QList != null) {
									str[7] = Integer.toString(QList.size());
								} else {
									str[7] = "0";
								}
								PSW = HDList.size();
								str[9] = Integer.toString(PSW);
								hsnum = hsnum + PSW;
							}
//							if (HDList != null) {
//								RSW = HDList.size();
//								str[8] = Integer.toString(RSW);
//								hsnum = hsnum + RSW;
//								rswnum = rswnum + RSW;
//							} else {
//								RSW = 0;
//								str[7] = Integer.toString(RSW);
//							}
							rownum++;

							list.add(str);
						}
						// 如果上料工序没有对应的点焊工序，也需要输出，不需要统计焊点
						if (oplist != null && oplist.size() > 0) {
							for (int n = 0; n < oplist.size(); n++) {
								int LARGE_PARTS = 0;// LARGE PARTS列值
								int SUPERLARGE_PARTS = 0;// SUPERLARGE_PARTS列值
								int MID_PARTS = 0;// MID PARTS列值
								int SMALL_PARTS = 0;// SMALL PARTS列值
								int CLAMP = 0;// CLAMP& UN-CLAMP列值

								String[] str = new String[10];// 共10个数据
								TCComponentBOMLine chil = (TCComponentBOMLine) oplist.get(n);
								String discretename = chil.getProperty("bl_rev_object_name");

								// 根据焊装工位查询零件
								ArrayList part = new ArrayList();// 零件集合

								getSolItmPart(chil, part);

								for (int k = 0; k < part.size(); k++) {
									TCComponentItemRevision rev = (TCComponentItemRevision) part.get(k);
									String partno = Util.getProperty(rev, "dfl9_part_no");
									if (partno.length() > 5) {
										partno = partno.substring(0, 5);
									}
									if (rule.containsKey(partno)) {
										String type = rule.get(partno);
										if (type.equals("SUPER LARGE PARTS")) {
											SUPERLARGE_PARTS++;
										}
										if (type.equals("LARGE PARTS")) {
											LARGE_PARTS++;
										}
										if (type.equals("MID PARTS")) {
											MID_PARTS++;
										}
										if (type.equals("SMALL PARTS")) {
											SMALL_PARTS++;
										}
									}
								}

								PARTS_NAME = plinename + " " + bl.getProperty("bl_rev_object_name") + "(" + discretename
										+ ")";

								str[0] = Integer.toString(rownum + 1);// 序号

								str[1] = PARTS_NAME;

								// SUPERLARGE_PARTS = 0;
								if (SUPERLARGE_PARTS != 0) {
									str[2] = Integer.toString(SUPERLARGE_PARTS);
								}
								// LARGE_PARTS = 0;
								if (LARGE_PARTS != 0) {
									str[3] = Integer.toString(LARGE_PARTS);
								}
								// MID_PARTS = 0;
								if (MID_PARTS != 0) {
									str[4] = Integer.toString(MID_PARTS);
								}
								// SMALL_PARTS = 0;
								if (SMALL_PARTS != 0) {
									str[5] = Integer.toString(SMALL_PARTS);
								}
								if (CLAMP != 0) {
									str[6] = Integer.toString(CLAMP);
								}
								list.add(str);
							}
						}
					}
					// 工位中无机器人工序
					else {
						int LARGE_PARTS = 0;// LARGE PARTS列值
						int SUPERLARGE_PARTS = 0;// SUPERLARGE_PARTS列值
						int MID_PARTS = 0;// MID PARTS列值
						int SMALL_PARTS = 0;// SMALL PARTS列值
						int CLAMP = 0;// CLAMP& UN-CLAMP列值

						PARTS_NAME = plinename + " " + bl.getProperty("bl_rev_object_name");

						// 根据焊装工位查询零件
						ArrayList part = new ArrayList();// 零件集合

						getSolItmPart(bl, part);

						for (int j = 0; j < part.size(); j++) {
							TCComponentItemRevision rev = (TCComponentItemRevision) part.get(j);
							String partno = Util.getProperty(rev, "dfl9_part_no");
							if (partno.length() > 5) {
								partno = partno.substring(0, 5);
							}
							if (rule.containsKey(partno)) {
								String type = rule.get(partno);
								if (type.equals("SUPER LARGE PARTS")) {
									SUPERLARGE_PARTS++;
								}
								if (type.equals("LARGE PARTS")) {
									LARGE_PARTS++;
								}
								if (type.equals("MID PARTS")) {
									MID_PARTS++;
								}
								if (type.equals("SMALL PARTS")) {
									SMALL_PARTS++;
								}
							}
						}

						String[] str = new String[10];// 共10个数据
						str[0] = Integer.toString(rownum + 1);// 序号
						str[1] = PARTS_NAME;

						// SUPERLARGE_PARTS = 0;
						if (SUPERLARGE_PARTS != 0) {
							str[2] = Integer.toString(SUPERLARGE_PARTS);
						}
						// LARGE_PARTS = 0;
						if (LARGE_PARTS != 0) {
							str[3] = Integer.toString(LARGE_PARTS);
						}
						// MID_PARTS = 0;
						if (MID_PARTS != 0) {
							str[4] = Integer.toString(MID_PARTS);
						}
						// SMALL_PARTS = 0;
						if (SMALL_PARTS != 0) {
							str[5] = Integer.toString(SMALL_PARTS);
						}
						if (CLAMP != 0) {
							str[6] = Integer.toString(CLAMP);
						}
						// CLAMP = 0;

//						ArrayList QList = Util.searchBOMLine(bl, "OR",
//								new String[] { "bl_item_object_type", "bl_item_object_type" }, "==",
//								new String[] { "枪", "BIW Gun" });
//						ArrayList QList = getGunNums(bl);
//						ArrayList HDList = Util.searchBOMLine(bl, "OR",
//								new String[] { "bl_item_object_type", "bl_item_object_type" }, "==",
//								new String[] { "焊点", "Weld Point" });
						Integer[] hdnums = getWeldPointNums(bl);
//						if (QList != null) {
//							GUN = QList.size();
//							str[7] = Integer.toString(GUN);
//						} else {
//							str[7] = "0";
//						}
						str[7] = getGunNums(bl).toString();
						if (hdnums != null) {
							PSW = hdnums[0];
							RSW = hdnums[1];
							if (RSW != 0) {
								str[8] = hdnums[1].toString();
							}
							if (PSW != 0) {
								str[9] = hdnums[0].toString();
							}
							hsnum = hsnum + PSW + RSW;
							rswnum = rswnum + RSW;
						}
						rownum++;
						list.add(str);
					}

				}

				// 装备焊
				if (ajm == 2) {
					int LARGE_PARTS = 0;// LARGE PARTS列值
					int SUPERLARGE_PARTS = 0;// SUPERLARGE_PARTS列值
					int MID_PARTS = 0;// MID PARTS列值
					int SMALL_PARTS = 0;// SMALL PARTS列值
					int CLAMP = 0;// CLAMP& UN-CLAMP列值

					PARTS_NAME = "FIX - " + bl.getProperty("bl_rev_object_name");

					// 根据焊装工位查询零件
					ArrayList part = new ArrayList();// 零件集合

					getSolItmPart(bl, part);

					for (int j = 0; j < part.size(); j++) {
						TCComponentItemRevision rev = (TCComponentItemRevision) part.get(j);
						String partno = rev.getProperty("dfl9_part_no");
						if (partno.length() > 5) {
							partno = partno.substring(0, 5);
						}
						if (rule.containsKey(partno)) {
							String type = rule.get(partno);
							if (type.equals("SUPER LARGE PARTS")) {
								SUPERLARGE_PARTS++;
							}
							if (type.equals("LARGE PARTS")) {
								LARGE_PARTS++;
							}
							if (type.equals("MID PARTS")) {
								MID_PARTS++;
							}
							if (type.equals("SMALL PARTS")) {
								SMALL_PARTS++;
							}
						}
					}

					String[] str = new String[10];// 共10个数据
					str[0] = Integer.toString(rownum + 1);// 序号
					str[1] = PARTS_NAME;

					// SUPERLARGE_PARTS = 0;
					if (SUPERLARGE_PARTS != 0) {
						str[2] = Integer.toString(SUPERLARGE_PARTS);
					}
					// LARGE_PARTS = 0;
					if (LARGE_PARTS != 0) {
						str[3] = Integer.toString(LARGE_PARTS);
					}
					// MID_PARTS = 0;
					if (MID_PARTS != 0) {
						str[4] = Integer.toString(MID_PARTS);
					}
					// SMALL_PARTS = 0;
					if (SMALL_PARTS != 0) {
						str[5] = Integer.toString(SMALL_PARTS);
					}
					if (CLAMP != 0) {
						str[6] = Integer.toString(CLAMP);
					}
					rownum++;

					list.add(str);
				}
			}
			// 装备焊不增加check等3行
			if (ajm == 2) {
				list.add(rownum);
			} else {
				list.add(rownum + 3);
			}

			return list;
		} catch (Exception ex) {
			ex.printStackTrace();
		}
		return list;
	}

	private Integer getGunNums(TCComponentBOMLine bl) throws TCException {
		// TODO Auto-generated method stub
		Integer gunnum = 0;
		AIFComponentContext[] children = bl.getChildren();
		for (AIFComponentContext chil : children) {
			TCComponentBOMLine disbl = (TCComponentBOMLine) chil.getComponent();
			TCComponentItemRevision rev = disbl.getItemRevision();
			if (rev.isTypeOf("B8_BIWDiscreteOPRevision")) {
				ArrayList gunlist = Util.getChildrenByBOMLine(disbl, "B8_BIWGunRevision");	
				if(gunlist!=null) {
					gunnum = gunnum + gunlist.size();
				}
			}
		}

		return gunnum;
	}

	/*
	 * 获取工位下的机器焊点和人工焊点的数量
	 */
	private Integer[] getWeldPointNums(TCComponentBOMLine bl) throws TCException {
		Integer[] nums = new Integer[2];
		int pswnum = 0;
		int rswnum = 0;
		AIFComponentContext[] children = bl.getChildren();
		for (AIFComponentContext chil : children) {
			TCComponentBOMLine disbl = (TCComponentBOMLine) chil.getComponent();
			TCComponentItemRevision rev = disbl.getItemRevision();
			if (rev.isTypeOf("B8_BIWDiscreteOPRevision")) {
				ArrayList weldlist = Util.getChildrenByBOMLine(disbl, "WeldPointRevision");
				int factnum = 0;
				if (weldlist != null) {
					factnum = weldlist.size();
				}
				String objectname = Util.getProperty(rev, "object_name");
				if (objectname.substring(0, 1).equals("R")) {
					rswnum = rswnum + factnum;
				} else {
					pswnum = pswnum + factnum;
				}
			}
		}
		nums[0] = pswnum;
		nums[1] = rswnum;

		return nums;
	}

	// 把METAL产线放到最后输出
	private ArrayList getOrderList(List CList) {
		// 排好序的产线集合
		ArrayList orderList = new ArrayList();
		// METAL产线
		ArrayList MList = new ArrayList();
		// 非METAL产线
		ArrayList UNMList = new ArrayList();

		for (int i = 0; i < CList.size(); i++) {
			TCComponentBOMLine bl = (TCComponentBOMLine) CList.get(i);
			int n = getIsAJ(bl);
			if (n == 2) {
				MList.add(bl);
			} else {
				UNMList.add(bl);
			}
		}
		for (int j = 0; j < UNMList.size(); j++) {
			orderList.add(UNMList.get(j));
		}
		for (int k = 0; k < MList.size(); k++) {
			orderList.add(MList.get(k));
		}

		return orderList;
	}

	private int getIsAJ(TCComponentBOMLine bl) {
		// TODO Auto-generated method stub
		int type = 1;
		try {
			String objectname = bl.getProperty("bl_rev_object_name");

			if ((objectname.contains("04") && objectname.contains("FM"))
					|| (objectname.contains("07") && objectname.contains("BM"))) {
				type = 1; // 需要区分机器人工序
			} else if (objectname.contains("METAL")) {
				type = 2; // METAl产线
			} else {
				type = 3;// 其他
			}
		} catch (TCException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		return type;
	}

	// 判断焊装工位工艺下是否有多个点焊工序
	private boolean getIsDiscretes(TCComponentBOMLine bl) {
		// TODO Auto-generated method stub
		boolean flag = false;
		try {
			AIFComponentContext[] children = bl.getChildren();

			for (AIFComponentContext chil : children) {
				TCComponentItemRevision rev = ((TCComponentBOMLine) chil.getComponent()).getItemRevision();
				if (rev.isTypeOf("B8_BIWDiscreteOPRevision") || rev.isTypeOf("B8_BIWOperationRevision")) {
					String objectname = Util.getProperty(rev, "object_name");
					if (objectname.substring(0, 1).equals("R")) {
						flag = true;
					}
				}
			}
		} catch (Exception e) {
			e.printStackTrace();
		}
		return flag;
	}

	private boolean getIsRSW(TCComponentBOMLine bl) {
		// TODO Auto-generated method stub
		boolean flag = false;
		try {
			AIFComponentContext[] children = bl.getChildren();
			int num = 0;
			for (AIFComponentContext chil : children) {
				TCComponentBOMLine downbl = (TCComponentBOMLine) chil.getComponent();
				TCComponentItemRevision rev = downbl.getItemRevision();
				if (rev.isTypeOf("B8_BIWDiscreteOPRevision")) {
					if (downbl.getProperty("bl_rev_object_name").substring(0, 1).equals("R")) {
						flag = true;
						break;
					}
				}
			}
			if (num > 1) {
				return true;
			}
		} catch (Exception e) {
			e.printStackTrace();
		}
		return flag;
	}

	// 通过焊装工位查询有多少零件DFL9SolItmPartRevision
	private ArrayList getSolItmPart(TCComponentBOMLine bl, ArrayList part) {
		try {
			AIFComponentContext[] children = bl.getChildren();
			int num = 0;
			for (AIFComponentContext chil : children) {
				TCComponentBOMLine ch = (TCComponentBOMLine) chil.getComponent();
				TCComponentItemRevision rev = ch.getItemRevision();
				if (rev.isTypeOf("DFL9SolItmPartRevision")) {
					part.add(rev);
				} else {
					getSolItmPart(ch, part);
				}
			}
			return part;
		} catch (Exception e) {
			e.printStackTrace();
		}
		return part;
	}

	// 查询大小件规则首选项，获取大小件规则信息
	private HashMap<String, String> getSizeRule() {
		HashMap<String, String> rule = new HashMap<String, String>();
		try {

			File file = null;
			Workbook workbook = null;
			TCSession session = (TCSession) app.getSession();
			TCPreferenceService preferenceService = session.getPreferenceService();

			String str = preferenceService.getPreferenceDescription("DFL9_straight_sheet_size_rule");
			if (str != null) {
				String value = preferenceService.getStringValue("DFL9_straight_sheet_size_rule");
				if (value != null) {
					TCComponentDatasetType datatype = (TCComponentDatasetType) session.getTypeComponent("Dataset");
					TCComponentDataset dataset = datatype.find(value);
					if (dataset != null) {
						String type = dataset.getType();

						TCComponentTcFile[] files;
						try {
							files = dataset.getTcFiles();
							if (files.length > 0) {
								file = files[0].getFmsFile();
							}
						} catch (TCException e) {
							// TODO Auto-generated catch block
							e.printStackTrace();
						}

						if (file != null) {
							FileInputStream inputStream = new FileInputStream(file);
							if (type.equals("MSExcel")) {
								workbook = new HSSFWorkbook(inputStream);
								rule = parseCoverExcel(workbook);
							}
							if (type.equals("MSExcelX")) {
								workbook = new XSSFWorkbook(inputStream);
								rule = parseCoverExcel(workbook);
							}
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

	private static HashMap<String, String> parseCoverExcel(Workbook workbook) {
		// TODO Auto-generated method stub
		HashMap<String, String> rule = new HashMap<String, String>();
		// 解析sheet

		Sheet sheet = workbook.getSheetAt(0);
		// 校验sheet是否合法
		if (sheet == null) {
			return null;
		}
		// 获取第一行数据
		int firstRowNum = sheet.getFirstRowNum();
		Row firstRow = (Row) sheet.getRow(firstRowNum);
		if (null == firstRow) {
			logger.warning("解析Excel失败，在第一行没有读取到任何数据！");
		}

		// 解析每一行的数据，构造数据对象
		int rowStart = firstRowNum + 1;
		int rowEnd = sheet.getPhysicalNumberOfRows();
		for (int rowNum = rowStart; rowNum < rowEnd; rowNum++) {
			Row row = (Row) sheet.getRow(rowNum);
			if (null == row) {
				continue;
			}
			String[] resultData = convertRowToCoverData(row);
			if (null == resultData) {
				logger.warning("第 " + row.getRowNum() + "行数据不合法，已忽略！");
				continue;
			}
			if (resultData[0] != null && !resultData[0].isEmpty()) {
				rule.put(resultData[0], resultData[1]);
			}
		}

		return rule;
	}

	private static String[] convertRowToCoverData(Row row) {
		// TODO Auto-generated method stub
		String[] value = new String[2];
		Cell cell;
		// 板件前5位
		cell = row.getCell(8);
		if (cell != null) {
			String partno = convertCellValueToString(cell, Cell.CELL_TYPE_STRING);
			value[0] = partno.trim();
		}
		// 板件类型
		cell = row.getCell(10);
		if (cell != null) {
			String parttype = convertCellValueToString(cell, Cell.CELL_TYPE_STRING);
			value[1] = parttype.trim();
		}
		return value;
	}

	private static String convertCellValueToString(Cell cell, int type) {
		if (cell == null) {
			return null;
		}
		String returnValue = null;
		switch (type) {
		case Cell.CELL_TYPE_NUMERIC: // 数字
			Double doubleValue = cell.getNumericCellValue();
			// 格式化科学计数法，取一位整数
			DecimalFormat df = new DecimalFormat("0.00");
			returnValue = df.format(doubleValue);
			break;
		case Cell.CELL_TYPE_STRING: // 字符串
			cell.setCellType(Cell.CELL_TYPE_STRING);
			returnValue = cell.getStringCellValue();
			break;
		case Cell.CELL_TYPE_BOOLEAN: // 布尔
			Boolean booleanValue = cell.getBooleanCellValue();
			returnValue = booleanValue.toString();
			break;
		case Cell.CELL_TYPE_BLANK: // 空值
			break;
		case Cell.CELL_TYPE_FORMULA: // 公式
			returnValue = cell.getCellFormula();
			break;
		case Cell.CELL_TYPE_ERROR: // 故障
			break;
		default:
			break;
		}
		return returnValue;
	}
}
