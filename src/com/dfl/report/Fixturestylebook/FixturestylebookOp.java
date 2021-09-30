
package com.dfl.report.Fixturestylebook;

import java.awt.image.BufferedImage;
import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
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
import java.util.LinkedList;
import java.util.List;
import java.util.Map;
import java.util.Map.Entry;

import javax.imageio.ImageIO;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.PrintSetup;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.ShapeTypes;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFAnchor;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFClientAnchor;
import org.apache.poi.xssf.usermodel.XSSFDrawing;
import org.apache.poi.xssf.usermodel.XSSFPrintSetup;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFShape;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFSimpleShape;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openxmlformats.schemas.drawingml.x2006.spreadsheetDrawing.CTShape;

import com.dfl.report.ExcelReader.BoardInformation;
import com.dfl.report.ExcelReader.WeldPointBoardInformation;
import com.dfl.report.ExcelReader.WeldPointInfo;
import com.dfl.report.common.TCComponentUtils;
import com.dfl.report.util.FileUtil;
import com.dfl.report.util.NewOutputDataToExcel;
import com.dfl.report.util.ReportViwePanel;
import com.dfl.report.util.Util;
import com.spire.pdf.PdfDocument;
import com.teamcenter.rac.aif.AbstractAIFUIApplication;
import com.teamcenter.rac.aif.kernel.AIFComponentContext;
import com.teamcenter.rac.aif.kernel.InterfaceAIFComponent;
import com.teamcenter.rac.aifrcp.AIFUtility;
import com.teamcenter.rac.cme.kernel.bvr.FlowUtil;
import com.teamcenter.rac.cme.kernel.mfg.IMfgFlow;
import com.teamcenter.rac.cme.kernel.mfg.IMfgNode;
import com.teamcenter.rac.kernel.DeepCopyInfo;
import com.teamcenter.rac.kernel.TCComponent;
import com.teamcenter.rac.kernel.TCComponentBOMLine;
import com.teamcenter.rac.kernel.TCComponentDataset;
import com.teamcenter.rac.kernel.TCComponentDatasetType;
import com.teamcenter.rac.kernel.TCComponentFolder;
import com.teamcenter.rac.kernel.TCComponentGroup;
import com.teamcenter.rac.kernel.TCComponentItem;
import com.teamcenter.rac.kernel.TCComponentItemRevision;
import com.teamcenter.rac.kernel.TCComponentItemType;
import com.teamcenter.rac.kernel.TCComponentProject;
import com.teamcenter.rac.kernel.TCComponentTcFile;
import com.teamcenter.rac.kernel.TCComponentUser;
import com.teamcenter.rac.kernel.TCException;
import com.teamcenter.rac.kernel.TCPreferenceService;
import com.teamcenter.rac.kernel.TCProperty;
import com.teamcenter.rac.kernel.TCSession;

public class FixturestylebookOp {

	private TCSession session;
	private InterfaceAIFComponent[] aifComponents;
	// private String page1;
	private String page2;
	SimpleDateFormat df = new SimpleDateFormat("yyyy/MM/dd");// 设置日期格式
	private LinkedHashMap<String, String> fymap = new LinkedHashMap<String, String>();// 用于部品数据分页

	private boolean IsUpdate;
	private ArrayList gwlist;
	// List<TCComponentProject> projects = new LinkedList<TCComponentProject>();
	TCComponent[] projects;
	private List tempPartlist = new ArrayList();
	private String isupdateflag;

	public FixturestylebookOp(TCSession session, InterfaceAIFComponent[] aifComponents, String page2, boolean IsUpdate,
			ArrayList list, String isupdateflag) throws TCException, AccessException {
		// TODO Auto-generated constructor stub
		this.session = session;
		this.aifComponents = aifComponents;
		// this.page1 = page1;
		this.page2 = page2;
		this.IsUpdate = IsUpdate;
		this.gwlist = list;
		this.isupdateflag = isupdateflag;
		initUI();
	}

	private void initUI() throws TCException, AccessException {
		// TODO Auto-generated method stub

		// 界面显示进度并输出执行步骤
		ReportViwePanel viewPanel = new ReportViwePanel("生成报表");
		viewPanel.setVisible(true);
		viewPanel.addInfomation("开始输出报表...\n", 10, 100);
		// 公共信息
		TCComponentBOMLine targrt = (TCComponentBOMLine) aifComponents[0];
		TCComponentBOMLine topbl = targrt.window().getTopBOMLine();// BOP顶层
		projects = topbl.getItemRevision().getRelatedComponents("project_list");
		TCComponentGroup group = session.getGroup();
		// 科室
		String groupname = group.getFullName();	
		String groupname1 = "";
		System.out.println("用户组：" + group.getLocalizedFullName());
		// 部门
		String department = "";
		if (groupname != null && (groupname.contains("新车准备技术部") || groupname.contains("New Model Preparation Engineering Department"))) {
			department = "新车准备技术部";
		} else if (groupname != null && (groupname.contains("车辆工程技术部") || groupname.contains("Vehicle Process Engineering Department"))) {
			department = "车辆工程技术部";
		} else {
			department = "";
		}
		if (groupname != null
				&& (groupname.contains("同期工程科") || groupname.contains("simultaneous Engineering Section"))) {
			groupname1 = "同期工程科";
		} else if (groupname != null
				&& (groupname.contains("焊装技术科") || groupname.contains("Body Assembly Engineering Section"))) {
			groupname1 = "焊装技术科";
		} else {
			groupname1 = "";
		}
		// 操作人
		TCComponentUser user = session.getUser();
		String username = Util.getProperty(user, "user_name");
		// 车型
		String VehicleNo = "";
		String project_ids = Util.getProperty(topbl, "bl_rev_project_ids");
		VehicleNo = Util.getDFLProjectIdVehicle(project_ids);

		for (int n = 0; n < gwlist.size(); n++) {
			double schedule = (n + 1.0) / gwlist.size();
			int sch = (int) (schedule * 100);
			if (sch < 100) {
				viewPanel.addInfomation("", sch, 100);
			}
			boolean multflag = false; // 是否多个工位
			boolean RLflag = false; // 是否左右工位
			TCComponentItemRevision oldrev = null;
			Map<String, File> Sectionlist = new HashMap<String, File>();

			tempPartlist.clear();

			TCComponentBOMLine gwbl = (TCComponentBOMLine) gwlist.get(n);// 工位
			TCComponentBOMLine linebl = gwbl.parent();// 产线
			String[] baseinfo = new String[14];
			baseinfo[0] = groupname1;
			baseinfo[1] = department;
			baseinfo[2] = username;
			baseinfo[3] = VehicleNo;
			// 生产线名称 如果产线下只有一个工位，为产线名称，去除产线名称中的LH或RH ;如果产线下有多个工位，为产线名称+工位名称，去除产线名称中的LH或RH
			String linename = "";
			String ProcLine = Util.getProperty(linebl, "bl_rev_object_name");
			String gwname = Util.getProperty(gwbl, "bl_rev_object_name");
			// 增加纤体式样
			String ProcLinetype = Util.getProperty(linebl, "bl_B8_BIWMEProcLineRevision_b8_LineType");

			baseinfo[12] = gwname;
			baseinfo[13] = ProcLine;
			ArrayList gwlist = Util.getChildrenByBOMLine(linebl, "B8_BIWMEProcStatRevision");
			// 根据工位的产线名称是否为LH还是RH
			boolean LRflag = false;
			if (ProcLine.length() > 1) {
				String rl = ProcLine.substring(ProcLine.length() - 2, ProcLine.length());
				if (rl.equals("LH") || rl.equals("RH")) {
					ProcLine = ProcLine.substring(0, ProcLine.length() - 2);
				}
				if (rl.equals("LH")) {
					LRflag = false;
				} else if (rl.equals("RH")) {
					LRflag = true;
				} else {
					LRflag = false;
				}
			}
			if (gwlist != null && gwlist.size() > 1) {
				linename = ProcLinetype + ProcLine + " " + gwname;
				multflag = true;
			} else {
				linename = ProcLinetype + ProcLine;
			}
			baseinfo[4] = linename;
			// 对称工位
			TCComponentBOMLine ssgwbl = getSymmetryState(linebl, gwname);

			// 部品番号
			String assyno = "";
			TCProperty p = gwbl.getItemRevision().getTCProperty("b8_ProcAssyNo2");

			String[] assy = null;
			String[] assy2 = null;
			if (ssgwbl != null) {
				TCProperty p2 = ssgwbl.getItemRevision().getTCProperty("b8_ProcAssyNo2");
				if (p2 != null) {
					assy2 = p2.getStringValueArray();
				}
			}
			if (p != null) {
				assy = p.getStringValueArray();
			}
			if (assy != null && assy.length > 0) {
				if (ssgwbl != null && assy[0].length() > 4) {
					if (assy2 != null && assy2.length > 0) {
						if (assy2[0].length() > 4) {
							String endno = assy2[0].substring(4, 5);
							assyno = assy[0].substring(0, 5) + "/" + endno + assy[0].substring(5, assy[0].length());
						} else {
							assyno = assy[0];
						}
					} else {
						assyno = assy[0];
					}

				} else {
					assyno = assy[0];
				}
			}
			baseinfo[11] = assyno;
			// 获取assy号
			List assylist = new ArrayList();
			List assynamelist = new ArrayList();
			// 获取assy号和名称
			if (assy != null) {
				for (int i = 0; i < assy.length; i++) {
					String[] str = new String[2];
					str[0] = assy[i];
					str[1] = Util.getProperty(linebl, "bl_rev_object_name");
					assylist.add(assy[i]);
					assynamelist.add(str);
				}
			}
			if (assy2 != null) {
				for (int i = 0; i < assy2.length; i++) {
					String ssprocline = Util.getProperty(ssgwbl.parent(), "bl_rev_object_name");
					String[] str = new String[2];
					str[0] = assy2[i];
					str[1] = ssprocline;
					assylist.add(assy2[i]);
					assynamelist.add(str);
				}
			}

			// 部品名称
			String assyname = ProcLine;
			baseinfo[5] = assyname;
			// 日期
			String date = df.format(new Date());
			baseinfo[6] = date;

			InputStream inputStream = null;

			// 根据选择的焊装工位工艺下，是否已经生成过报表，如果生成过则直接取之前的报表作为模板
			TCComponentItemRevision blrev = gwbl.getItemRevision();
			// 输出的文件名称
			String datasetname = gwname + "夹具式样书";
			String filename = Util.formatString(datasetname);
			TCComponent[] tccs = blrev.getRelatedComponents("IMAN_reference");
			TCComponentItem tcc = null;
			System.out.println("关系对象数组：" + tccs);
			for (TCComponent item : tccs) {
				String type = Util.getRelProperty(item, "object_name");
				if (type.equals(datasetname)) {
					tcc = (TCComponentItem) item;
					break;
				}
			}
			System.out.println("关系对象：" + tcc);
			if (tcc != null) {
				oldrev = tcc.getLatestItemRevision();
			}
			if (IsUpdate) {
				if (oldrev == null) {
					System.out.println("获取文件失败，请确认是否已生成过报表！");
					// viewPanel.addInfomation("获取文件失败，请确认是否已生成过报表！\n", 100, 100);
					return;
				}
				TCComponent[] tccdata = oldrev.getRelatedComponents("IMAN_specification");
				TCComponentDataset dataset = null;
				File file = null;
				if (tccdata != null && tccdata.length > 0) {
					dataset = (TCComponentDataset) tccdata[0];
				}
				if (dataset != null) {
					file = dataset.getFile("excel", filename + ".xlsx", dataset.getWorkingDir());
				} else {
					System.out.println("获取数据集失败！");
				}

				// 先删除图片
//				File scripfile = Util.getRCPPluginInsideFile("RemovePicture.vbs");
//				if(scripfile!=null) {
//					String filepath = file.getPath();
//					String command = "wscript " + scripfile.getAbsolutePath() + " \"" +filepath+"\" ";
//					System.out.println("command:" + command);
//					
//					try {
//						Process process = Runtime.getRuntime().exec(command);
//						process.waitFor();
//					} catch (IOException | InterruptedException e) {
//						// TODO Auto-generated catch block
//						e.printStackTrace();
//					}
//				}
				if (file == null) {
					System.out.println("获取文件失败，请确认是否有修改过文件名称或者文档对象下存在报表数据集！");
					inputStream = FileUtil.getTemplateFile("DFL_Template_FixtureStyleBook");
					if (inputStream == null) {
						System.out.println("错误：没有找到夹具式样书模板，请先添加模板(名称为：DFL_Template_FixtureStyleBook)");
					}
				} else {
					try {
						inputStream = new FileInputStream(file);
					} catch (FileNotFoundException e) {
						// TODO Auto-generated catch block
						e.printStackTrace();
					}
				}

			} else {
				// 查询目录导出模板
				inputStream = FileUtil.getTemplateFile("DFL_Template_FixtureStyleBook");

				if (inputStream == null) {
					System.out.println("错误：没有找到夹具式样书模板，请先添加模板(名称为：DFL_Template_FixtureStyleBook)");
				}

			}

			// 夹具数量
			String jjnum = "";
//			String typename = Util.getObjectDisplayName(session, "B8_BIWFixture");
//			String[] propertys2 = new String[] { "bl_item_object_type", "bl_item_object_type" };
//			String[] values2 = new String[] { typename, typename };
//			ArrayList jjlist = Util.searchBOMLine(gwbl, "OR", propertys2, "==", values2);
//			if (jjlist != null && jjlist.size() > 0) {
//				jjnum = Integer.toString(jjlist.size());
//			} else {
//				jjnum = "0";
//			}
			baseinfo[7] = jjnum + "台";
			// GUN形式、GAUGE板厚
			String[] guninfo = getGunInfomation(gwbl);
			baseinfo[8] = guninfo[0];
			baseinfo[9] = guninfo[1];
			baseinfo[10] = guninfo[2];

			// 获取部品信息
			List RHlist = getPartsinformation(gwbl);

			List LHlist = new ArrayList();

			if (ssgwbl != null) {
				RLflag = true;
				LHlist = getPartsinformation(ssgwbl);
			}

			// 设置标号并排序
			ArrayList partlist = SetLabelsAndSort(RHlist, gwbl, ssgwbl, LHlist);// 部品数据集

			// getRLHStateData(sortList, LHlist, partlist);

			// 获取Locate List信息
			List Lglllist = getDatumGLLInfo(gwbl);
			List Rglllist = new ArrayList();
			if (ssgwbl != null) {
				// RHllist = getDatumGLLInfo(ssgwbl);
				Rglllist = getRHGLLinfo(Lglllist, LRflag);
			}
			// 获取枪的集合
			List gunlist = getGunInfo(gwbl);

			// 获取部品构成图
			Map<String, File> Rhpiclist = new HashMap<String, File>();
			Map<String, File> Lhpiclist = getAll3DPictures(blrev, "1");
			if (ssgwbl != null) {
				Rhpiclist = getAll3DPictures(ssgwbl.getItemRevision(), "1");
			}
			// 断面定位形状仕様
			Sectionlist = getAll3DPictures(blrev, "2");

			// viewPanel.addInfomation("开始写数据，请耐心等待...\n", 60, 100);

			XSSFWorkbook book = creatXSSFWorkbook(inputStream, page2, gwname, RLflag, partlist);

			// 写入部品构成图
			if (ssgwbl != null) {
				writePartCharDataToSheet(book, Lhpiclist, LRflag);
				writePartCharDataToSheet(book, Rhpiclist, !LRflag);
			} else {
				writePartCharDataToSheet(book, Lhpiclist, false);
			}
			if (IsUpdate) 
			{
				if("3".equals(isupdateflag))
				{
					// 写构成部品一覧
					writePartDataToSheet(book, baseinfo, assynamelist, partlist);
				}
			}
			else
			{
				// 写构成部品一覧
				writePartDataToSheet(book, baseinfo, assynamelist, partlist);
			}
			

			// 写仕様差一覧
			writePoorPatternProcessing(book, baseinfo, assylist, RLflag, partlist);
			// 左工位放在左边，右工位放在右边
			if (ssgwbl != null) {
				if (LRflag) {
					writeLocateListDataToSheet(book, baseinfo, Rglllist, Lglllist);
				} else {
					writeLocateListDataToSheet(book, baseinfo, Lglllist, Rglllist);
				}
			} else {
				// Locate List
				writeLocateListDataToSheet(book, baseinfo, Lglllist, Rglllist);
			}

			// 写断面定位形状仕様
			writeSectionDataToSheet(book, Sectionlist);

			// 写STD GUN Drawing
			writeSTDGUNDataToSheet(book, gunlist, baseinfo);

			// 如果是更新需要判断是否增加页数
			if (IsUpdate) {
				// writeOtherpages(book, "断面定位形状仕様");
				writeOtherpages(book, "weld layout");
			}
			// 对sheet重命名
			SetSheetRename(book, gwname, multflag);

			// 写入主数据
			writeMainDataToSheet(book, baseinfo);

			NewOutputDataToExcel.exportFile(book, filename);
			// viewPanel.addInfomation("", 80, 100);

			saveFiles(datasetname, filename, gwbl, ssgwbl, oldrev);

		}
		viewPanel.addInfomation("输出报表完成，请在各焊装工位工艺版本下附件查看！", 100, 100);

	}

	/*
	 * 写断面定位形状仕様
	 */
	private void writeSectionDataToSheet(XSSFWorkbook book, Map<String, File> Sectionlist) {
		// TODO Auto-generated method stub
		if (Sectionlist != null && Sectionlist.size() > 0) {
			int sheetnum = 0;
			sheetnum = book.getNumberOfSheets();
			int sheetAtIndex = -1; //
			for (int i = 0; i < sheetnum; i++) {
				String sheetname = book.getSheetName(i);
				if (sheetname.contains("9_断面定位形状仕様")) {
					sheetAtIndex = i;
					break;
				}
			}
			if (sheetAtIndex == -1) {
				return;
			}

			// 根据数据判断是否需要分页
			int page = (Sectionlist.size() - 1) / 6;

			int index = sheetAtIndex + 1;

			/**************************************************/
			// 如果是更新，则先把系统输出的信息清空，后面再写入
			if (IsUpdate) {
				int gcnum = 0;
				for (int i = 0; i < sheetnum; i++) {
					String sheetname = book.getSheetName(i);
					if (sheetname.contains("9_断面定位形状仕様")) {
						gcnum++;
					}
				}
				// 如果sheet页增加就增，减少不删除，保留
				index = gcnum + sheetAtIndex;

				// 循环构成表sheet页清空系统输出内容，手工维护内容保留
				for (int i = sheetAtIndex; i < index; i++) {
					// 循环遍历获取图片名称，再匹配写入

				}
				if (gcnum < page) {
					for (int i = 0; i < page - gcnum; i++) {
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
			int count = 0;
			int rowindex = 0;
			int colindex = 0;
			int shnum = 0;
			XSSFSheet sheet = null;
			if (IsUpdate) {
				// 更新前数据总数
				int predatanum = Sectionlist.size();
				// 循环构成表sheet页根据名称更新图片
				for (int i = sheetAtIndex; i < index; i++) {
					// 循环遍历获取图片名称，再匹配写入
					XSSFSheet sh = book.getSheetAt(sheetAtIndex);
					String picname = parseExcel(sh, 4, 2);
					updatePicDataTosheet(book, sh, picname, null, 0, 0, Sectionlist);
					String picname1 = parseExcel(sh, 21, 2);
					updatePicDataTosheet(book, sh, picname1, null, 1, 0, Sectionlist);
					String picname2 = parseExcel(sh, 38, 2);
					updatePicDataTosheet(book, sh, picname2, null, 2, 0, Sectionlist);
					String picname3 = parseExcel(sh, 4, 41);
					updatePicDataTosheet(book, sh, picname3, null, 0, 1, Sectionlist);
					String picname4 = parseExcel(sh, 21, 41);
					updatePicDataTosheet(book, sh, picname4, null, 1, 1, Sectionlist);
					String picname5 = parseExcel(sh, 38, 41);
					updatePicDataTosheet(book, sh, picname5, null, 2, 1, Sectionlist);
				}
				// 更新后数据总数
				int afterdatanum = Sectionlist.size();
				// 把新增的数据写入
				if (afterdatanum > 0) {
					// 先计算从哪个位置写入
					int diff = predatanum - afterdatanum;
					// 第几个sheet
					int shAt = (diff - 1) / 6;
					// 新增的从第几个开始
					int Dataindex = diff % 6;
					// 写入之前sheet页空留的位置数据
					if (Dataindex != 0) {
						for (Entry<String, File> entry : Sectionlist.entrySet()) {
							if (Dataindex > 5) {
								if ((Dataindex + 1) % 3 == 1) {
									rowindex = 0;
								} else if ((Dataindex + 1) % 3 == 2) {
									rowindex = 1;
								} else {
									rowindex = 2;
								}
								colindex = (Dataindex) / 3;
								Dataindex++;
								colindex = (count) / 3;
								sheet = book.getSheetAt(shAt);
								count++;
								String objectname = entry.getKey();
								File file = entry.getValue();
								String[] values = objectname.split(" ");
								if (values != null && values.length > 0) {
									for (int i = 0; i < values.length; i++) {
										setStringCellAndStyle(sheet, values[i], 4 + i * 2 + 17 * rowindex,
												2 + 39 * colindex, null, Cell.CELL_TYPE_STRING);//
									}
								}

								BufferedImage image = null;
								try {
									image = ImageIO.read(file);
								} catch (IOException e) {
									// TODO Auto-generated catch block
									e.printStackTrace();
								}
								writepicturetosheet(book, sheet, image, 5 + rowindex * 17, 21 + colindex * 39,
										13 + rowindex * 17, 33 + colindex * 39);
							}
						}
					}
					// 剩下的从新的sheet页写入
					if (Sectionlist.size() > 0) {
						count = 0;
						rowindex = 0;
						colindex = 0;
						sheetAtIndex = shAt + 1;
						for (Entry<String, File> entry : Sectionlist.entrySet()) {
							if ((count + 1) % 3 == 1) {
								rowindex = 0;
							} else if ((count + 1) % 3 == 2) {
								rowindex = 1;
							} else {
								rowindex = 2;
							}
							colindex = (count) / 3;
							if (count % 6 == 0) {
								sheet = book.getSheetAt(sheetAtIndex + shnum);
								shnum++;
							}
							count++;
							String objectname = entry.getKey();
							File file = entry.getValue();
							String[] values = objectname.split(" ");
							if (values != null && values.length > 0) {
								for (int i = 0; i < values.length; i++) {
									setStringCellAndStyle(sheet, values[i], 4 + i * 2 + 17 * rowindex,
											2 + 39 * colindex, null, Cell.CELL_TYPE_STRING);// Ａｓｓｙ名称
								}
							}

							BufferedImage image = null;
							try {
								image = ImageIO.read(file);
							} catch (IOException e) {
								// TODO Auto-generated catch block
								e.printStackTrace();
							}
							writepicturetosheet(book, sheet, image, 5 + rowindex * 17, 21 + colindex * 39,
									13 + rowindex * 17, 33 + colindex * 39);

						}
					}

				}

			} else {
				for (Entry<String, File> entry : Sectionlist.entrySet()) {
					if ((count + 1) % 3 == 1) {
						rowindex = 0;
					} else if ((count + 1) % 3 == 2) {
						rowindex = 1;
					} else {
						rowindex = 2;
					}
					colindex = (count) / 3;
					if (count % 6 == 0) {
						sheet = book.getSheetAt(sheetAtIndex + shnum);
						shnum++;
					}
					count++;
					String objectname = entry.getKey();
					File file = entry.getValue();
					String[] values = objectname.split(" ");
					if (values != null && values.length > 0) {
						for (int i = 0; i < values.length; i++) {
							setStringCellAndStyle(sheet, values[i], 4 + i * 2 + 17 * rowindex, 2 + 39 * colindex, null,
									Cell.CELL_TYPE_STRING);// Ａｓｓｙ名称
						}
					}

					BufferedImage image = null;
					try {
						image = ImageIO.read(file);
					} catch (IOException e) {
						// TODO Auto-generated catch block
						e.printStackTrace();
					}
					writepicturetosheet(book, sheet, image, 5 + rowindex * 17, 21 + colindex * 39, 13 + rowindex * 17,
							33 + colindex * 39);

				}
			}

		}
	}

	/*
	 * 更新断面定位形状仕様
	 */
	private void updatePicDataTosheet(XSSFWorkbook book, XSSFSheet sheet, String picname, XSSFCellStyle style,
			int rowindex, int colindex, Map<String, File> Sectionlist) {
		if (Sectionlist.containsKey(picname)) {
			String objectname = picname;
			File file = Sectionlist.get(picname);
			String[] values = objectname.split(" ");
			if (values != null && values.length > 0) {
				for (int i = 0; i < values.length; i++) {
					setStringCellAndStyle(sheet, values[i], 4 + i * 2 + 17 * rowindex, 2 + 39 * colindex, style,
							Cell.CELL_TYPE_STRING);// Ａｓｓｙ名称
				}
			}

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
			int h = 8;
			int w = (int) (h * diff);

			writepicturetosheet(book, sheet, image, 5 + rowindex * 17, 21 + colindex * 39, 13 + rowindex * 17,
					(21 + w) + colindex * 39);

			Sectionlist.remove(picname);
		}
	}

	/*
	 * 写入部品构成图
	 */
	private void writePartCharDataToSheet(XSSFWorkbook book, Map<String, File> lhpiclist, boolean flag) {
		// TODO Auto-generated method stub
		if (lhpiclist != null && lhpiclist.size() > 0) {
			int sheetnum = 0;
			sheetnum = book.getNumberOfSheets();
			int sheetAtIndex = -1; //
			for (int i = 0; i < sheetnum; i++) {
				String sheetname = book.getSheetName(i);
				if (sheetname.contains("4_部品构成图")) {
					sheetAtIndex = i;
					break;
				}
			}
			if (sheetAtIndex == -1) {
				return;
			}
			if (flag) {
				sheetAtIndex++;
			}
			// 设置字体颜色
			Font font = book.createFont();
			font.setColor((short) 2);//
			font.setBoldweight(Font.BOLDWEIGHT_BOLD); // 字体加粗
			font.setFontHeightInPoints((short) 14);

			XSSFSheet sheet = book.getSheetAt(sheetAtIndex);

			// 画图的顶级管理器对象HSSFPatriarch, 一个sheet只能获取一个
			XSSFDrawing hssfPatriarch = (XSSFDrawing) sheet.createDrawingPatriarch();
			XSSFClientAnchor anchor1 = null;
			XSSFRichTextString strValue = new XSSFRichTextString();
			int count = 0;
			int rowindex = 0;
			int colindex = 0;

			for (Entry<String, File> entry : lhpiclist.entrySet()) {
				if ((count + 1) % 3 == 1) {
					rowindex = 0;
				} else if ((count + 1) % 3 == 2) {
					rowindex = 1;
				} else {
					rowindex = 2;
				}
				colindex = (count) / 3;
				count++;
				String objectname = entry.getKey().replace("/", "  ");
				File file = entry.getValue();
				anchor1 = new XSSFClientAnchor(-80000, -80000, 80000, 80000, (short) (5 + colindex * 19),
						4 + rowindex * 14, (short) (15 + colindex * 19), 6 + rowindex * 14);
				// 创建一个矩形
				if (anchor1 != null) {
					anchor1.setAnchorType(2);
					XSSFSimpleShape rect = hssfPatriarch.createSimpleShape(anchor1);
					rect.setShapeType(ShapeTypes.RECT);
					rect.setLineStyleColor(0, 0, 0);
					rect.setLineWidth(0.75);
					rect.setNoFill(false);
					strValue.setString(objectname);
					strValue.applyFont(font);
					rect.setText(strValue);
				}
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
				int h = 10;
				int w = (int) (10 * diff);

				writepicturetosheet(book, sheet, image, 7 + rowindex * 14, 3 + colindex * 19, 17 + rowindex * 14,
						(3 + w) + colindex * 19);
			}

		}
	}

	private void writeOtherpages(XSSFWorkbook book, String pagename) {
		// TODO Auto-generated method stub
		int sheetnum = 0;
		sheetnum = book.getNumberOfSheets();
		int sheetAtIndex = -1; //
		for (int i = 0; i < sheetnum; i++) {
			String sheetname = book.getSheetName(i);
			if (sheetname.contains(pagename)) {
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
			if (sheetname.contains(pagename)) {
				gcnum++;
			}
		}
		// 如果sheet页增加就增，减少不删除，保留
		int index = sheetAtIndex + gcnum;
		int pagenum = 0;
		if (pagename.equals("weld layout")) {
			pagenum = Integer.parseInt(page2);
		}
		if (gcnum < pagenum) {
			for (int i = 0; i < pagenum - gcnum; i++) {
				XSSFSheet newsheet = book.cloneSheet(sheetAtIndex);
				book.setSheetOrder(newsheet.getSheetName(), index);
				index++;
			}
		}

	}

	/*
	 * 对sheet重命名
	 */
	private void SetSheetRename(XSSFWorkbook book, String gwname, boolean multflag) {
		int sheetnum = book.getNumberOfSheets();
		int aa = 0;// 部品构成图
		int bb = 0;// 构成部品一覧
		int cc = 0;// 仕様差一覧
		int dd = 0;// Gauge layout
		int ee = 0;// 断面定位形状仕様
		int ff = 0;// weld layout
		int gg = 0;// STD GUN Drawing
		int aaIndex = 0;// 部品构成图
		int bbIndex = 0;// 构成部品一覧
		int ccIndex = 0;// 仕様差一覧
		int ddIndex = 0;// Gauge layout
		int eeIndex = 0;// 断面定位形状仕様
		int ffIndex = 0;// weld layout
		int ggIndex = 0;// STD GUN Drawing
		for (int i = 0; i < sheetnum; i++) {
			XSSFSheet sheet = book.getSheetAt(i);
			String sheetname = sheet.getSheetName();
			if (multflag && sheetname.contains("基本仕様")) {
				book.setSheetName(i, "P3_基本仕様" + gwname);
			}
			if (sheetname.contains("部品构成图")) {
				aa++;
				aaIndex = i;
			}
			if (sheetname.contains("构成部品一覧")) {
				bb++;
				bbIndex = i;
			}
			if (sheetname.contains("仕様差一覧")) {
				cc++;
				ccIndex = i;
			}
			if (sheetname.contains("Gauge layout")) {
				dd++;
				ddIndex = i;
			}
			if (sheetname.contains("断面定位形状仕様")) {
				ee++;
				eeIndex = i;
			}
			if (sheetname.contains("weld layout")) {
				ff++;
				ffIndex = i;
			}
			if (sheetname.contains("STD GUN Drawing")) {
				gg++;
				ggIndex = i;
			}
		}
		if (aa > 1) {
			String aftername = "";
			if (multflag) {
				aftername = gwname + "(RH)";
			} else {
				aftername = "(RH)";
			}
			book.setSheetName(aaIndex, "4_部品构成图" + aftername);

			if (multflag) {
				aftername = gwname + "(LH)";
			} else {
				aftername = "(LH)";
			}
			book.setSheetName(aaIndex - 1, "4_部品构成图" + aftername);

		} else {
			if (multflag) {
				// System.out.println("4_部品构成图sheet重明明：" + aaIndex);
				book.setSheetName(aaIndex, "4_部品构成图" + gwname);
			}
		}
		if (dd > 1) {
			String aftername = "";
			aftername = "(RH)";
			book.setSheetName(ddIndex, "7_Gauge layout" + aftername);
			aftername = "(LH)";
			System.out.println("7_Gauge layout页数:" + ddIndex);
			book.setSheetName(ddIndex - 1, "7_Gauge layout" + aftername);

		}
		Integer[] inter = { bb, cc, ee, ff, gg };
		Integer[] interIndex = { bbIndex, ccIndex, eeIndex, ffIndex, ggIndex };
		String[] basename = { "5_构成部品一覧", "6.仕様差一覧", "9_断面定位形状仕様", "10_weld layout", "14_STD GUN Drawing" };
		for (int i = 0; i < inter.length; i++) {
			int sheetnums = inter[i];
			int index = interIndex[i];
			if (sheetnums > 1) {
				for (int j = 0; j < sheetnums; j++) {
					String sheetname = basename[i];
					String newsheetname = sheetname + "_" + Integer.toString((sheetnums - j));
					book.setSheetName((index - j), newsheetname);
				}
			}
		}

	}

	/*
	 * 写仕様差一覧
	 */
	private void writePoorPatternProcessing(XSSFWorkbook book, String[] baseinfo, List assylist, boolean rLflag2,
			ArrayList partlist) {
		// TODO Auto-generated method stub
		ArrayList poorlist = new ArrayList();
		// 获取式样差sheet
		int sheetnum = 0;
		sheetnum = book.getNumberOfSheets();
		int sheetAtIndex = -1; // 式样差所在位置
		for (int i = 0; i < sheetnum; i++) {
			String sheetname = book.getSheetName(i);
			if (sheetname.contains("仕様差一覧")) {
				sheetAtIndex = i;
				break;
			}
		}
		if (sheetAtIndex == -1) {
			return;
		}
		int poornum = 0;// 式样差件数量
		for (Map.Entry<String, String> entry : fymap.entrySet()) {
			String key = entry.getKey();
			int value = Integer.parseInt(entry.getValue());
			if (value > 1) {
				poornum++;
				List temp = new ArrayList();
				List afterName = new ArrayList(); // 部品番号后5位
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
				// 如果是左右工位，需要左右工位一行输出
				if (rLflag2) {
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
							String[] str3 = new String[3];
							if (str1[0].length() > 4 && str2[0].length() > 4) {
								str3[0] = str1[0].substring(0, 5) + "/" + str2[0].substring(4, 5) + ApartNO;
							} else {
								str3[0] = str1[0];
							}
							if (str1[1] != null && str1[1].length() > 1) {
								str3[1] = str1[1].substring(0, str1[1].length() - 2) + "L/RH";
							} else {
								str3[1] = str1[1] + "L/RH";
							}

							str3[2] = str1[2];
							poorlist.add(str3);
						} else {
							if (ttlist.size() > 0) {
								String[] str1 = (String[]) ttlist.get(0);
								poorlist.add(str1);
							}
						}
					}

				} else {
					for (int j = 0; j < temp.size(); j++) {
						String[] val = (String[]) temp.get(j);
						poorlist.add(val);
					}
				}
			}
		}
		// 根据数据判断是否需要分页,每4个行数据分一页
		int page = poornum / 4 + 1;

		// 数据行刚好一页就会出现sheet页多了一页的情况
		if (poornum % 4 == 0) {
			if (page > 1) {
				page = page - 1;
			}
		}

		// 如果page大于1，则需要复制sheet页
		int index = sheetAtIndex + 1;

		/**************************************************/
		// 如果是更新，则先把系统输出的信息清空，后面再写入
		if (IsUpdate) {
			int gcnum = 0;
			for (int i = 0; i < sheetnum; i++) {
				String sheetname = book.getSheetName(i);
				if (sheetname.contains("仕様差一覧")) {
					gcnum++;
				}
			}
			// 如果sheet页增加就增，减少不删除，保留
			index = sheetAtIndex + gcnum;

			// 循环构成表sheet页清空系统输出内容，手工维护内容保留
			for (int i = sheetAtIndex; i < index; i++) {
				XSSFSheet sheet = book.getSheetAt(i);
				// 清空内容
				setStringCellAndStyle(sheet, "", 7, 2, null, Cell.CELL_TYPE_STRING);// Ａｓｓｙ名称
				setStringCellAndStyle(sheet, "", 7, 20, null, Cell.CELL_TYPE_STRING);// Ａｓｓｙ名称
				setStringCellAndStyle(sheet, "", 7, 38, null, Cell.CELL_TYPE_STRING);// Ａｓｓｙ名称
				setStringCellAndStyle(sheet, "", 7, 56, null, Cell.CELL_TYPE_STRING);// Ａｓｓｙ名称

				for (int j = 0; j < 4; j++) {
					setStringCellAndStyle(sheet, "", 31 + j * 2, 2, null, Cell.CELL_TYPE_STRING);// Ａｓｓｙ部番
					setStringCellAndStyle(sheet, "", 31 + j * 2, 20, null, Cell.CELL_TYPE_STRING);// Ａｓｓｙ部番
					setStringCellAndStyle(sheet, "", 31 + j * 2, 38, null, Cell.CELL_TYPE_STRING);// Ａｓｓｙ部番
					setStringCellAndStyle(sheet, "", 31 + j * 2, 56, null, Cell.CELL_TYPE_STRING);// Ａｓｓｙ部番

				}
			}
			if (gcnum < page) {
				for (int i = 0; i < page - gcnum; i++) {
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

			// 写公共部分
			setStringCellAndStyle(sheet, baseinfo[3], 2, 1, null, Cell.CELL_TYPE_STRING);// 车型
			setStringCellAndStyle(sheet, baseinfo[2], 2, 66, null, Cell.CELL_TYPE_STRING);// 编制
			setStringCellAndStyle(sheet, baseinfo[6], 2, 72, null, Cell.CELL_TYPE_STRING);// 日期
			setStringCellAndStyle(sheet, baseinfo[3], 51, 41, null, Cell.CELL_TYPE_STRING);// 车型
			setStringCellAndStyle(sheet, baseinfo[4], 53, 44, null, Cell.CELL_TYPE_STRING);// 生产线
			setStringCellAndStyle(sheet, baseinfo[11], 53, 61, null, Cell.CELL_TYPE_STRING);// 工程番号

			if (i == index - 1) {
				for (int j = 0; j + 4 * shnum < poornum; j++) {
					String partname = "";
					int rownum = 0;
					for (int k = 0; k < poorlist.size(); k++) {
						String[] str = (String[]) poorlist.get(k);
						if (j + 1 + 4 * shnum == Integer.parseInt(str[2])) {
							partname = str[1];
							if ((j + 1 + 4 * shnum) % 4 == 1) {
								setStringCellAndStyle(sheet, str[0], 31 + rownum * 2, 2, null, Cell.CELL_TYPE_STRING);// Ａｓｓｙ部番
							} else if ((j + 1 + 4 * shnum) % 4 == 2) {
								setStringCellAndStyle(sheet, str[0], 31 + rownum * 2, 20, null, Cell.CELL_TYPE_STRING);// Ａｓｓｙ部番
							} else if ((j + 1 + 4 * shnum) % 4 == 3) {
								setStringCellAndStyle(sheet, str[0], 31 + rownum * 2, 38, null, Cell.CELL_TYPE_STRING);// Ａｓｓｙ部番
							} else {
								setStringCellAndStyle(sheet, str[0], 31 + rownum * 2, 56, null, Cell.CELL_TYPE_STRING);// Ａｓｓｙ部番
							}
							rownum++;
						}
					}
					if ((j + 1 + 4 * shnum) % 4 == 1) {
						setStringCellAndStyle(sheet, partname, 7, 2, null, Cell.CELL_TYPE_STRING);// Ａｓｓｙ名称
					} else if ((j + 1 + 4 * shnum) % 4 == 2) {
						setStringCellAndStyle(sheet, partname, 7, 20, null, Cell.CELL_TYPE_STRING);// Ａｓｓｙ名称
					} else if ((j + 1 + 4 * shnum) % 4 == 3) {
						setStringCellAndStyle(sheet, partname, 7, 38, null, Cell.CELL_TYPE_STRING);// Ａｓｓｙ名称
					} else {
						setStringCellAndStyle(sheet, partname, 7, 56, null, Cell.CELL_TYPE_STRING);// Ａｓｓｙ名称
					}

				}
			} else {
				for (int j = 0; j + 4 * shnum < 4 + 4 * shnum; j++) {
					String partname = "";
					int rownum = 0;
					for (int k = 0; k < poorlist.size(); k++) {
						String[] str = (String[]) poorlist.get(k);
						if (j + 1 + 4 * shnum == Integer.parseInt(str[2])) {
							partname = str[1];
							if ((j + 1 + 4 * shnum) % 4 == 1) {
								setStringCellAndStyle(sheet, str[0], 31 + rownum * 2, 2, null, Cell.CELL_TYPE_STRING);// Ａｓｓｙ部番
							} else if ((j + 1 + 4 * shnum) % 4 == 2) {
								setStringCellAndStyle(sheet, str[0], 31 + rownum * 2, 20, null, Cell.CELL_TYPE_STRING);// Ａｓｓｙ部番
							} else if ((j + 1 + 4 * shnum) % 4 == 3) {
								setStringCellAndStyle(sheet, str[0], 31 + rownum * 2, 38, null, Cell.CELL_TYPE_STRING);// Ａｓｓｙ部番
							} else {
								setStringCellAndStyle(sheet, str[0], 31 + rownum * 2, 56, null, Cell.CELL_TYPE_STRING);// Ａｓｓｙ部番
							}
							rownum++;
						}
					}
					if ((j + 1 + 4 * shnum) % 4 == 1) {
						setStringCellAndStyle(sheet, partname, 7, 2, null, Cell.CELL_TYPE_STRING);// Ａｓｓｙ名称
					} else if ((j + 1 + 4 * shnum) % 4 == 2) {
						setStringCellAndStyle(sheet, partname, 7, 20, null, Cell.CELL_TYPE_STRING);// Ａｓｓｙ名称
					} else if ((j + 1 + 4 * shnum) % 4 == 3) {
						setStringCellAndStyle(sheet, partname, 7, 38, null, Cell.CELL_TYPE_STRING);// Ａｓｓｙ名称
					} else {
						setStringCellAndStyle(sheet, partname, 7, 56, null, Cell.CELL_TYPE_STRING);// Ａｓｓｙ名称
					}
				}
			}
			shnum++;
		}
	}

	/*
	 * 写STD GUN Drawing
	 */
	private void writeSTDGUNDataToSheet(XSSFWorkbook book, List gunlist, String[] baseinfo) throws TCException {
		// TODO Auto-generated method stub

		int sheetnum = 0;
		sheetnum = book.getNumberOfSheets();
		int sheetAtIndex = -1; // Locate List所在位置
		for (int i = 0; i < sheetnum; i++) {
			String sheetname = book.getSheetName(i);
			if (sheetname.contains("STD GUN Drawing")) {
				sheetAtIndex = i;
				break;
			}
		}
		if (sheetAtIndex == -1) {
			return;
		}
		if (gunlist == null || gunlist.size() < 1) {
			XSSFSheet sheet = book.getSheetAt(sheetAtIndex);
			setStringCellAndStyle(sheet, baseinfo[3], 48, 35, null, Cell.CELL_TYPE_STRING);// 车型
			setStringCellAndStyle(sheet, baseinfo[4], 50, 35, null, Cell.CELL_TYPE_STRING);// 生产线
			return;
		}
		// 根据数据判断是否需要分页
		int page = gunlist.size();

		int index = sheetAtIndex + 1;

		/**************************************************/
		// 如果是更新，则先把系统输出的信息清空，后面再写入
		if (IsUpdate) {
			int gcnum = 0;
			for (int i = 0; i < sheetnum; i++) {
				String sheetname = book.getSheetName(i);
				if (sheetname.contains("STD GUN Drawing")) {
					gcnum++;
				}
			}
			// 如果sheet页增加就增，减少不删除，保留
			// index = sheetAtIndex + page;
			index = index + gcnum - 1 ;

			// 循环构成表sheet页清空系统输出内容，手工维护内容保留
			for (int i = sheetAtIndex; i < index; i++) {
				
			}
			if (gcnum < page) {
				for (int i = 0; i < page - gcnum; i++) {
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

		int count = 0;
		for (int i = sheetAtIndex; i < index; i++) {
			XSSFSheet sheet = book.getSheetAt(i);
			TCComponentBOMLine bl = (TCComponentBOMLine) gunlist.get(count);
			String gunname = Util.getProperty(bl, "bl_B8_BIWGunRevision_b8_Model");

			setStringCellAndStyle(sheet, baseinfo[3], 48, 35, null, Cell.CELL_TYPE_STRING);// 车型
			setStringCellAndStyle(sheet, baseinfo[4], 50, 35, null, Cell.CELL_TYPE_STRING);// 生产线
			setStringCellAndStyle(sheet, gunname, 50, 49, null, Cell.CELL_TYPE_STRING);// 枪

			BufferedImage image = getCalculationParameter(session, bl.getItemRevision());
			if (image != null) {
				writepicturetosheet(book, sheet, image, 2, 3, 46, 54);
			}
			count++;
		}

	}

	// 根据单个文件写图片到excel
	private static void writepicturetosheet(XSSFWorkbook book, XSSFSheet sheet, BufferedImage bufferImg, int rowindex,
			int colindex, int rowindex2, int colindex2) {
		// 先把读进来的图片放到一个ByteArrayOutputStream中，以便产生ByteArray
		ByteArrayOutputStream byteArrayOut = new ByteArrayOutputStream();
		try {
			ImageIO.write(bufferImg, "png", byteArrayOut);
			XSSFDrawing patriarch = sheet.createDrawingPatriarch();
			XSSFClientAnchor anchor = new XSSFClientAnchor(50000, 50000, -50000, -50000, (short) colindex, rowindex,
					(short) colindex2, rowindex2);
			anchor.setAnchorType(2);
			// 插入图片
			patriarch.createPicture(anchor,
					book.addPicture(byteArrayOut.toByteArray(), XSSFWorkbook.PICTURE_TYPE_JPEG));
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}

	}

	/*
	 * ****************************** 获取PDF截图
	 */
	public static BufferedImage getCalculationParameter(TCSession session, TCComponentItemRevision rev) {
		BufferedImage obj = null;
		try {
			File file = null;
			TCComponentDataset basicdata = null;
			TCComponent[] tccs = rev.getRelatedComponents("IMAN_specification");
			for (TCComponent item : tccs) {
				if (item instanceof TCComponentDataset) {
					String type = item.getType();
					if (type.equals("PDF")) {
						basicdata = (TCComponentDataset) item;
					}
				}
			}
			if (basicdata != null) {
				String type = basicdata.getType();
				if (type.equals("PDF")) {
					TCComponentTcFile[] files;
					try {
						files = basicdata.getTcFiles();
						if (files.length > 0) {
							file = files[0].getFmsFile();
						}
					} catch (TCException e) {
						// TODO Auto-generated catch block
						e.printStackTrace();
					}
				}
			}
			if (file != null) {
				PdfDocument doc = new PdfDocument();
				doc.loadFromFile(file.getPath());
				// 把第一页作为截图
				BufferedImage pageimge = doc.saveAsImage(0);
				obj = pageimge;
			}
			return obj;
		} catch (Exception e) {
			e.printStackTrace();
		}
		return obj;
	}

	/*
	 * 枪集合
	 */
	private List getGunInfo(TCComponentBOMLine gwbl) {
		// TODO Auto-generated method stub
		List gun = new ArrayList();
//		String[] propertys = new String[] { "bl_item_object_type", "bl_item_object_type" };
//		String[] values = new String[] { "枪", "BIW Gun" };
//		gun = Util.searchBOMLine(gwbl, "OR", propertys, "==", values);
		try {
			AIFComponentContext[] chidrens = gwbl.getChildren();
			for (AIFComponentContext aif : chidrens) {
				TCComponentBOMLine direbl = (TCComponentBOMLine) aif.getComponent();
				System.out.println("用于卡在10%的问题打印：" + direbl);
				if (direbl.getItemRevision().isTypeOf("B8_BIWDiscreteOPRevision")) {
					AIFComponentContext[] chids = direbl.getChildren();
					for (AIFComponentContext aifgun : chids) {
						TCComponentBOMLine gunbl = (TCComponentBOMLine) aifgun.getComponent();
						if (gunbl.getItemRevision().isTypeOf("B8_BIWGunRevision")) {
							gun.add(gunbl);
						}
					}
				}
			}
		} catch (TCException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		return gun;
	}

	/*
	 * 获取制定区域文本框
	 */
	private XSSFSimpleShape getXSSFSimpleShape(XSSFSheet sheet, int col1, int row1, int col2, int row2) {

		XSSFSimpleShape shape = null;
		XSSFDrawing hssfPatriarch = (XSSFDrawing) sheet.createDrawingPatriarch();
		List<XSSFShape> shapes = hssfPatriarch.getShapes();
		if (shapes != null) {
			for (int i = 0; i < shapes.size(); i++) {
				XSSFShape ss = hssfPatriarch.getShapes().get(i);
				if (ss instanceof XSSFSimpleShape) {
					XSSFSimpleShape a = (XSSFSimpleShape) hssfPatriarch.getShapes().get(i);
					XSSFShape b = hssfPatriarch.getShapes().get(i);
					// System.out.println(ShapeTypes.RECT + " " + a.getShapeType());
					if (a.getShapeType() == ShapeTypes.RECT) {
						shape = a;
					}
//					XSSFClientAnchor anchor = (XSSFClientAnchor) a.getAnchor();
//					if (anchor == null) {
//						continue;
//					}
//					int factcol1 = anchor.getCol1();
//					int factcol2 = anchor.getCol2();
//					int factrow1 = anchor.getRow1();
//					int factrow2 = anchor.getRow2();
//					// 如果在指定范围内，返回文本框
//					if (col1 >= factcol1 && col2 <= factcol2 && row1 >= factrow1 && row2 >= factrow2) {
//						shape = a;
//					}
				}
			}
		}
		return shape;
	}

	/*
	 * Locate List
	 */
	private void writeLocateListDataToSheet(XSSFWorkbook book, String[] baseinfo, List lglllist, List rglllist) {
		// TODO Auto-generated method stub
		int sheetnum = 0;
		sheetnum = book.getNumberOfSheets();
		int sheetAtIndex = -1; // Locate List所在位置
		for (int i = 0; i < sheetnum; i++) {
			String sheetname = book.getSheetName(i);
			if (sheetname.contains("Locate List")) {
				sheetAtIndex = i;
				break;
			}
		}
		if (sheetAtIndex == -1) {
			return;
		}
		// 根据数据判断是否需要分页
		int sum = lglllist.size();

		// 每14行分一个sheet页
		int page = sum / 14 + 1;

		// 数据行刚好一页就会出现sheet页多了一页的情况
		if (sum % 14 == 0) {
			if (page > 1) {
				page = page - 1;
			}
		}
		int index = sheetAtIndex + 1;

		/**************************************************/
		// 如果是更新，则先把系统输出的信息清空，后面再写入
		if (IsUpdate) {
			int gcnum = 0;
			for (int i = 0; i < sheetnum; i++) {
				String sheetname = book.getSheetName(i);
				if (sheetname.contains("Locate List")) {
					gcnum++;
				}
			}
			// 如果sheet页增加就增，减少不删除，保留
			index = sheetAtIndex + gcnum;

			// 循环构成表sheet页清空系统输出内容，手工维护内容保留
			for (int i = sheetAtIndex; i < index; i++) {
				XSSFSheet sheet = book.getSheetAt(i);
				// 清空内容
				for (int j = 0; j < 6; j++) {
					setStringCellAndStyle(sheet, "", 6 + 3 * j, 1, null, Cell.CELL_TYPE_STRING);// LOCATENo
					setStringCellAndStyle(sheet, "", 6 + 3 * j, 3, null, Cell.CELL_TYPE_STRING);// partname
					setStringCellAndStyle(sheet, "", 7 + 3 * j, 3, null, Cell.CELL_TYPE_STRING);// partno
					setStringCellAndStyle(sheet, "", 6 + 3 * j, 13, null, Cell.CELL_TYPE_STRING);// HOLE SIZE
					setStringCellAndStyle(sheet, "", 6 + 3 * j, 14, null, Cell.CELL_TYPE_STRING);//
					setStringCellAndStyle(sheet, "", 6 + 3 * j, 16, null, Cell.CELL_TYPE_STRING);//
					setStringCellAndStyle(sheet, "", 6 + 3 * j, 17, null, Cell.CELL_TYPE_STRING);// PIN SIZE
					setStringCellAndStyle(sheet, "", 5 + 3 * j, 19, null, Cell.CELL_TYPE_STRING);//
					setStringCellAndStyle(sheet, "", 6 + 3 * j, 19, null, Cell.CELL_TYPE_STRING);// POSITION
					setStringCellAndStyle(sheet, "", 7 + 3 * j, 19, null, Cell.CELL_TYPE_STRING);//
					setStringCellAndStyle(sheet, "", 5 + 3 * j, 20, null, Cell.CELL_TYPE_STRING);//
					setStringCellAndStyle(sheet, "", 6 + 3 * j, 20, null, Cell.CELL_TYPE_STRING);// POSITION
					setStringCellAndStyle(sheet, "", 7 + 3 * j, 20, null, Cell.CELL_TYPE_STRING);//

					setStringCellAndStyle(sheet, "", 6 + 3 * j, 1, null, Cell.CELL_TYPE_STRING);// LOCATENo
					setStringCellAndStyle(sheet, "", 6 + 3 * j, 32, null, Cell.CELL_TYPE_STRING);// partno
					setStringCellAndStyle(sheet, "", 7 + 3 * j, 32, null, Cell.CELL_TYPE_STRING);// partname
					setStringCellAndStyle(sheet, "", 6 + 3 * j, 42, null, Cell.CELL_TYPE_STRING);// HOLE SIZE
					setStringCellAndStyle(sheet, "", 6 + 3 * j, 43, null, Cell.CELL_TYPE_STRING);//
					setStringCellAndStyle(sheet, "", 6 + 3 * j, 45, null, Cell.CELL_TYPE_STRING);//
					setStringCellAndStyle(sheet, "", 6 + 3 * j, 46, null, Cell.CELL_TYPE_STRING);// PIN SIZE
					setStringCellAndStyle(sheet, "", 5 + 3 * j, 48, null, Cell.CELL_TYPE_STRING);//
					setStringCellAndStyle(sheet, "", 6 + 3 * j, 48, null, Cell.CELL_TYPE_STRING);// POSITION
					setStringCellAndStyle(sheet, "", 7 + 3 * j, 48, null, Cell.CELL_TYPE_STRING);//
					setStringCellAndStyle(sheet, "", 5 + 3 * j, 49, null, Cell.CELL_TYPE_STRING);//
					setStringCellAndStyle(sheet, "", 6 + 3 * j, 49, null, Cell.CELL_TYPE_STRING);// POSITION
					setStringCellAndStyle(sheet, "", 7 + 3 * j, 49, null, Cell.CELL_TYPE_STRING);//
				}
			}
			if (gcnum < page) {
				for (int i = 0; i < page - gcnum; i++) {
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
		// 写Locate List数据
		XSSFRichTextString Richstr = new XSSFRichTextString();
		// 设置字体颜色
		Font font = book.createFont();
		font.setColor((short) 2);//
		font.setFontHeightInPoints((short) 11);

		int shnum = 0;

		for (int i = sheetAtIndex; i < index; i++) {
			XSSFSheet sheet = book.getSheetAt(i);
			// 画图的顶级管理器对象HSSFPatriarch, 一个sheet只能获取一个
			XSSFDrawing hssfPatriarch = (XSSFDrawing) sheet.createDrawingPatriarch();
			XSSFClientAnchor anchor1 = null;
			XSSFRichTextString strValue = new XSSFRichTextString();
			// 写公共部分
			setStringCellAndStyle(sheet, baseinfo[3], 48, 34, null, Cell.CELL_TYPE_STRING);// 车型
			setStringCellAndStyle(sheet, baseinfo[4], 50, 34, null, Cell.CELL_TYPE_STRING);// 生产线

			if (i == index - 1) {
				// 写LH工位信息
				for (int j = 0; j + 14 * shnum < lglllist.size(); j++) {
					String[] str = (String[]) lglllist.get(j + 14 * shnum);
					anchor1 = new XSSFClientAnchor(40000, -80000, 200000, 70000, (short) (1), 6 + 3 * j, (short) (2),
							7 + 3 * j);
					anchor1.setAnchorType(2);
					XSSFSimpleShape rect = hssfPatriarch.createSimpleShape(anchor1);
					rect.setShapeType(ShapeTypes.RECT);
					rect.setLineStyleColor(0, 0, 0);
					rect.setLineWidth(0.75);
					// rect.setNoFill(false);
					rect.setFillColor(255, 255, 255);// 白色背景色
					strValue.setString(str[12]);
					strValue.applyFont(font);
					rect.setText(strValue);
					// setStringCellAndStyle(sheet, str[12], 5 + 3*j, 1, style4,
					// Cell.CELL_TYPE_STRING);// LOCATENo
					setStringCellAndStyle(sheet, str[0], 7 + 3 * j, 3, null, Cell.CELL_TYPE_STRING);// partno
					setStringCellAndStyle(sheet, str[1], 6 + 3 * j, 3, null, Cell.CELL_TYPE_STRING);// partname
					setStringCellAndStyle(sheet, str[2], 6 + 3 * j, 13, null, Cell.CELL_TYPE_STRING);// HOLE SIZE
					setStringCellAndStyle(sheet, str[3], 6 + 3 * j, 14, null, Cell.CELL_TYPE_STRING);//
					setStringCellAndStyle(sheet, str[4], 6 + 3 * j, 16, null, Cell.CELL_TYPE_STRING);//
					setStringCellAndStyle(sheet, str[5], 6 + 3 * j, 17, null, Cell.CELL_TYPE_STRING);// PIN SIZE
					setStringCellAndStyle(sheet, str[6], 5 + 3 * j, 19, null, Cell.CELL_TYPE_STRING);//
					setStringCellAndStyle(sheet, str[7], 6 + 3 * j, 19, null, Cell.CELL_TYPE_STRING);// POSITION
					setStringCellAndStyle(sheet, str[8], 7 + 3 * j, 19, null, Cell.CELL_TYPE_STRING);//
					setStringCellAndStyle(sheet, str[9], 5 + 3 * j, 20, null, Cell.CELL_TYPE_STRING);//
					setStringCellAndStyle(sheet, str[10], 6 + 3 * j, 20, null, Cell.CELL_TYPE_STRING);// POSITION
					setStringCellAndStyle(sheet, str[11], 7 + 3 * j, 20, null, Cell.CELL_TYPE_STRING);//
				}
				if (rglllist != null && rglllist.size() > 0) {
					// 写RH工位信息
					for (int j = 0; j + 14 * shnum < rglllist.size(); j++) {
						String[] str = (String[]) rglllist.get(j + 14 * shnum);
						anchor1 = new XSSFClientAnchor(40000, -80000, 200000, 70000, (short) (30), 6 + 3 * j,
								(short) (31), 7 + 3 * j);
						anchor1.setAnchorType(2);
						XSSFSimpleShape rect = hssfPatriarch.createSimpleShape(anchor1);
						rect.setShapeType(ShapeTypes.RECT);
						rect.setLineStyleColor(0, 0, 0);
						rect.setLineWidth(0.75);
						// rect.setNoFill(false);
						rect.setFillColor(255, 255, 255);
						strValue.setString(str[12]);
						strValue.applyFont(font);
						rect.setText(strValue);
						// setStringCellAndStyle(sheet, str[12], 5 + 3*j, 1, style4,
						// Cell.CELL_TYPE_STRING);// LOCATENo
						setStringCellAndStyle(sheet, str[0], 7 + 3 * j, 32, null, Cell.CELL_TYPE_STRING);// partno
						setStringCellAndStyle(sheet, str[1], 6 + 3 * j, 32, null, Cell.CELL_TYPE_STRING);// partname
						setStringCellAndStyle(sheet, str[2], 6 + 3 * j, 42, null, Cell.CELL_TYPE_STRING);// HOLE SIZE
						setStringCellAndStyle(sheet, str[3], 6 + 3 * j, 43, null, Cell.CELL_TYPE_STRING);//
						setStringCellAndStyle(sheet, str[4], 6 + 3 * j, 45, null, Cell.CELL_TYPE_STRING);//
						setStringCellAndStyle(sheet, str[5], 6 + 3 * j, 46, null, Cell.CELL_TYPE_STRING);// PIN SIZE
						setStringCellAndStyle(sheet, str[6], 5 + 3 * j, 48, null, Cell.CELL_TYPE_STRING);//
						setStringCellAndStyle(sheet, str[7], 6 + 3 * j, 48, null, Cell.CELL_TYPE_STRING);// POSITION
						setStringCellAndStyle(sheet, str[8], 7 + 3 * j, 48, null, Cell.CELL_TYPE_STRING);//
						setStringCellAndStyle(sheet, str[9], 5 + 3 * j, 49, null, Cell.CELL_TYPE_STRING);//
						setStringCellAndStyle(sheet, str[10], 6 + 3 * j, 49, null, Cell.CELL_TYPE_STRING);// POSITION
						setStringCellAndStyle(sheet, str[11], 7 + 3 * j, 49, null, Cell.CELL_TYPE_STRING);//
					}
				}
			} else {
				// 写LH工位信息
				for (int j = 0; j + 14 * shnum < 14 + 14 * shnum; j++) {
					String[] str = (String[]) lglllist.get(j + 14 * shnum);
					anchor1 = new XSSFClientAnchor(100000, -80000, 180000, 70000, (short) (1), 6 + 3 * j, (short) (2),
							7 + 3 * j);
					anchor1.setAnchorType(2);
					XSSFSimpleShape rect = hssfPatriarch.createSimpleShape(anchor1);
					rect.setShapeType(ShapeTypes.RECT);
					rect.setLineStyleColor(0, 0, 0);
					rect.setLineWidth(0.75);
					// rect.setNoFill(false);
					rect.setFillColor(255, 255, 255);
					strValue.setString(str[12]);
					strValue.applyFont(font);
					rect.setText(strValue);
					// setStringCellAndStyle(sheet, str[12], 5 + 3*j, 1, style4,
					// Cell.CELL_TYPE_STRING);// LOCATENo
					setStringCellAndStyle(sheet, str[0], 7 + 3 * j, 3, null, Cell.CELL_TYPE_STRING);// partno
					setStringCellAndStyle(sheet, str[1], 6 + 3 * j, 3, null, Cell.CELL_TYPE_STRING);// partname
					setStringCellAndStyle(sheet, str[2], 6 + 3 * j, 13, null, Cell.CELL_TYPE_STRING);// HOLE SIZE
					setStringCellAndStyle(sheet, str[3], 6 + 3 * j, 14, null, Cell.CELL_TYPE_STRING);//
					setStringCellAndStyle(sheet, str[4], 6 + 3 * j, 16, null, Cell.CELL_TYPE_STRING);//
					setStringCellAndStyle(sheet, str[5], 6 + 3 * j, 17, null, Cell.CELL_TYPE_STRING);// PIN SIZE
					setStringCellAndStyle(sheet, str[6], 5 + 3 * j, 19, null, Cell.CELL_TYPE_STRING);//
					setStringCellAndStyle(sheet, str[7], 6 + 3 * j, 19, null, Cell.CELL_TYPE_STRING);// POSITION
					setStringCellAndStyle(sheet, str[8], 7 + 3 * j, 19, null, Cell.CELL_TYPE_STRING);//
					setStringCellAndStyle(sheet, str[9], 5 + 3 * j, 20, null, Cell.CELL_TYPE_STRING);//
					setStringCellAndStyle(sheet, str[10], 6 + 3 * j, 20, null, Cell.CELL_TYPE_STRING);// POSITION
					setStringCellAndStyle(sheet, str[11], 7 + 3 * j, 20, null, Cell.CELL_TYPE_STRING);//
				}
				if (rglllist != null && rglllist.size() > 0) {
					// 写RH工位信息
					for (int j = 0; j + 14 * shnum < 14 + 14 * shnum; j++) {
						String[] str = (String[]) rglllist.get(j + 14 * shnum);
						anchor1 = new XSSFClientAnchor(100000, -80000, 180000, 70000, (short) (30), 6 + 3 * j,
								(short) (31), 7 + 3 * j);
						anchor1.setAnchorType(2);
						XSSFSimpleShape rect = hssfPatriarch.createSimpleShape(anchor1);
						rect.setShapeType(ShapeTypes.RECT);
						rect.setLineStyleColor(0, 0, 0);
						rect.setLineWidth(0.75);
						// rect.setNoFill(false);
						rect.setFillColor(255, 255, 255);
						strValue.setString(str[12]);
						strValue.applyFont(font);
						rect.setText(strValue);
						// setStringCellAndStyle(sheet, str[12], 5 + 3*j, 1, style4,
						// Cell.CELL_TYPE_STRING);// LOCATENo
						setStringCellAndStyle(sheet, str[0], 7 + 3 * j, 32, null, Cell.CELL_TYPE_STRING);// partno
						setStringCellAndStyle(sheet, str[1], 6 + 3 * j, 32, null, Cell.CELL_TYPE_STRING);// partname
						setStringCellAndStyle(sheet, str[2], 6 + 3 * j, 42, null, Cell.CELL_TYPE_STRING);// HOLE SIZE
						setStringCellAndStyle(sheet, str[3], 6 + 3 * j, 43, null, Cell.CELL_TYPE_STRING);//
						setStringCellAndStyle(sheet, str[4], 6 + 3 * j, 45, null, Cell.CELL_TYPE_STRING);//
						setStringCellAndStyle(sheet, str[5], 6 + 3 * j, 46, null, Cell.CELL_TYPE_STRING);// PIN SIZE
						setStringCellAndStyle(sheet, str[6], 5 + 3 * j, 48, null, Cell.CELL_TYPE_STRING);//
						setStringCellAndStyle(sheet, str[7], 6 + 3 * j, 48, null, Cell.CELL_TYPE_STRING);// POSITION
						setStringCellAndStyle(sheet, str[8], 7 + 3 * j, 48, null, Cell.CELL_TYPE_STRING);//
						setStringCellAndStyle(sheet, str[9], 5 + 3 * j, 49, null, Cell.CELL_TYPE_STRING);//
						setStringCellAndStyle(sheet, str[10], 6 + 3 * j, 49, null, Cell.CELL_TYPE_STRING);// POSITION
						setStringCellAndStyle(sheet, str[11], 7 + 3 * j, 49, null, Cell.CELL_TYPE_STRING);//
					}
				}
			}
			shnum++;
		}

	}

	private List getRHGLLinfo(List lglllist, boolean lRflag) {
		// TODO Auto-generated method stub
		List RHList = new ArrayList();
		if (lglllist != null && lglllist.size() > 0) {
			for (int i = 0; i < lglllist.size(); i++) {
				String[] values = (String[]) lglllist.get(i);
				String[] rhval = new String[values.length];
				for (int j = 0; j < values.length; j++) {
					rhval[j] = values[j];
				}
				if (rhval[0] != null && rhval[0].length() > 5) {
					String partno = rhval[0].substring(4, 5);
					System.out.println("处理前的零件号：" + rhval[0]);
					if (Util.isNumber(partno)) {
						// 如果是左工位，需要减1，如果是右工位需要加1
						if (!lRflag) {
							int inno = Integer.parseInt(partno) - 1;
							String partNum = rhval[0].substring(0, 4) + Integer.toString(inno)
									+ rhval[0].substring(5, rhval[0].length());
							System.out.println("处理后的零件号：" + partNum);
							rhval[0] = partNum;
						} else {
							int inno = Integer.parseInt(partno) + 1;
							String partNum = rhval[0].substring(0, 4) + Integer.toString(inno)
									+ rhval[0].substring(5, rhval[0].length());
							System.out.println("处理后的零件号：" + partNum);
							rhval[0] = partNum;
						}

					}
				}
				if (rhval[1] != null && rhval[1].length() > 2) {
					System.out.println("处理前的零件名称：" + rhval[1]);
					String hz = rhval[1].substring(rhval[1].length() - 2);
					if (lRflag) {
						if (hz.equals("LH") || hz.equals("RH")) {
							rhval[1] = rhval[1].substring(0, rhval[1].length() - 2) + "LH";
						} else {
							rhval[1] = rhval[1] + "LH";
						}
					} else {
						if (hz.equals("LH") || hz.equals("RH")) {
							rhval[1] = rhval[1].substring(0, rhval[1].length() - 2) + "RH";
						} else {
							rhval[1] = rhval[1] + "RH";
						}
					}
					System.out.println("处理后的零件名称：" + rhval[1]);
				}
				if (rhval[10] != null && Util.isNumber(rhval[10])) {
					if (Double.parseDouble(rhval[10]) != 0) {
						rhval[10] = Double.toString(-1 * Double.parseDouble(rhval[10]));
					} else {
						rhval[10] = Double.toString(Double.parseDouble(rhval[10]));
					}

				}
				RHList.add(rhval);
			}
		}
		return RHList;
	}

	private List getDatumGLLInfo(TCComponentBOMLine gwbl) throws TCException {
		// TODO Auto-generated method stub

		List dglist = new ArrayList();
		String typename = Util.getObjectDisplayName(session, "B8_MPContainer");
		String[] propertys2 = new String[] { "bl_item_object_type", "bl_item_object_type" };
		String[] values2 = new String[] { typename, typename };
		ArrayList list = Util.searchBOMLine(gwbl, "OR", propertys2, "==", values2);
		if (list != null && list.size() > 0) {
			for (int i = 0; i < list.size(); i++) {
				TCComponentBOMLine bl = (TCComponentBOMLine) list.get(i);
				String objectname = Util.getProperty(bl, "bl_rev_object_name");
				if (objectname != null && !objectname.isEmpty()) {
					System.out.println(objectname.substring(0, 1));
					if (objectname.substring(0, 1).equals("L")) {
						// PARTS NAME(PARTS NO)
						String PARTS_NO = Util.getProperty(bl.getItemRevision(), "b8_ConnectedPartNo");
						String PARTS_NAME = getPropertysBypartNo(bl.window().getTopBOMLine(), PARTS_NO);
						String PinSize = Util.getProperty(bl, "B8_PinSize");
						ArrayList Datumlist = Util.getChildrenByBOMLine(bl, "DatumPointRevision");
						if (Datumlist != null && Datumlist.size() > 0) {
							for (int j = 0; j < Datumlist.size(); j++) {
								TCComponentBOMLine dbl = (TCComponentBOMLine) Datumlist.get(j);
								String[] values = new String[13];
								values[12] = objectname;
								values[0] = PARTS_NO;
								values[1] = PARTS_NAME;
								String datumname = Util.getProperty(dbl, "bl_rev_object_name");
								// 获取x,y,z坐标
								String xform = Util.getProperty(dbl, "bl_plmxml_abs_xform");// 绝对变换矩阵
								Double[] xyzArray = getXYZ(xform);
								Double x = xyzArray[0] * 1000;
								Double y = xyzArray[1] * 1000;
								Double z = xyzArray[2] * 1000;
								double diff = 0.2;
								if (PinSize != null && Util.isNumber(PinSize)) {
									diff = Double.parseDouble(PinSize);
								}
								// HOLE SIZE // PIN SIZE
								if (datumname.substring(0, 1).equals("C")) {
									String[] datumVal = datumname.split("[φΦ]");
									if (datumVal.length > 1) {
										values[2] = "";
										String[] tempdatum = datumVal[1].split("_");
										values[3] = tempdatum[0];
										String[] val = values[3].split("[×x]");
										System.out.println("转换后的数值：" + val[0]);
										if (Util.isNumber(val[0])) {
											values[5] = Double.toString(Double.parseDouble(val[0]) - diff);
										} else {
											values[5] = "";
										}
									} else {
										values[2] = "";
										values[3] = datumname;
										String val = values[3].substring(1);
										if (Util.isNumber(val)) {
											values[5] = Double.toString(Double.parseDouble(val) - diff);
										} else {
											values[5] = "";
										}
									}
								} else {
									values[2] = "φ";
									String[] trmpdatum = datumname.split("_");
									values[3] = trmpdatum[0].replace("φ", "");
									values[3] = values[3].replace("Φ", "");
									if (Util.isNumber(values[3]) && !values[3].contains("×")
											&& !values[3].contains("_")) {
										values[5] = Double.toString(Double.parseDouble(values[3]) - diff);
									} else {
										values[5] = "";
									}
								}
								values[4] = "φ";
								// POSITION
								values[6] = "X";
								values[7] = "Y";
								values[8] = "Z";
								values[9] = Double.toString(x);
								values[10] = Double.toString(y);
								values[11] = Double.toString(z);
								dglist.add(values);
							}
						}
					}
				}
			}
		}
		return dglist;
	}

	// 获取焊点的坐标（x,y,z）
	private Double[] getXYZ(String xform) {
		// TODO Auto-generated method stub
		Double[] values = new Double[] { 0.0, 0.0, 0.0 };
		String[] array = xform.split(" ");
		if (array != null && array.length == 16) {
			values[0] = Double.valueOf(array[12]);
			values[1] = Double.valueOf(array[13]);
			values[2] = Double.valueOf(array[14]);
		}
		return values;
	}

	// 调用查询获取板件属性
	private String getPropertysBypartNo(TCComponentBOMLine root, String partno) throws TCException {
		String values = "";
		// 调用系统查询，获取相关的板件
		List tcclist = Util.callStructureSearch(root, "__DFL_Find_SolutionPart", new String[] { "PARTNO" },
				new String[] { partno });
		if (tcclist != null && tcclist.size() > 0) {
			TCComponentBOMLine sol = (TCComponentBOMLine) tcclist.get(0);
			TCComponentItemRevision solrev3 = sol.getItemRevision();
			// values = Util.getProperty(solrev3, "object_name");// 板组
			values = Util.getProperty(solrev3, "dfl9_CADObjectName");// 板组
			System.out.println(partno + "板件名称：" + values);
		}
		return values;
	}

	/*
	 * 写构成部品一覧
	 */
	private void writePartDataToSheet(XSSFWorkbook book, String[] baseinfo, List assylist, ArrayList partlist) {
		// TODO Auto-generated method stub
		int sheetnum = 0;
		sheetnum = book.getNumberOfSheets();
		int sheetAtIndex = -1; // 构成表所在位置
		for (int i = 0; i < sheetnum; i++) {
			String sheetname = book.getSheetName(i);
			if (sheetname.contains("构成部品一覧")) {
				sheetAtIndex = i;
				break;
			}
		}
		if (sheetAtIndex == -1) {
			return;
		}
		// 根据数据判断是否需要分页
		int sum = 0;
//		for (Map.Entry<String, String> entry : fymap.entrySet()) {
//			sum = sum + Integer.parseInt(entry.getValue()) + 1;
//		}
		if (partlist != null) {
			sum = partlist.size();
		}
		// 每61行分一个sheet页
		int page = sum / 61 + 1;

		System.out.println("page:" + page);
		// 数据行刚好一页就会出现sheet页多了一页的情况
		if (sum % 61 == 0) {
			if (page > 1) {
				page = page - 1;
			}
		}
		int index = sheetAtIndex + 1;

		/***********************************************/
		// 如果是更新，则先把系统输出的信息清空，后面再写入
		if (IsUpdate) {
			int gcnum = 0;
			for (int i = 0; i < sheetnum; i++) {
				String sheetname = book.getSheetName(i);
				if (sheetname.contains("构成部品一覧")) {
					gcnum++;
				}
			}
			// 如果sheet页增加就增，减少不删除，保留
			index = sheetAtIndex + gcnum;

			// 循环构成表sheet页清空系统输出内容，手工维护内容保留
			for (int i = sheetAtIndex; i < index; i++) {
				XSSFSheet sheet = book.getSheetAt(i);
				// 清空partlist上部分信息
				for (int k = 0; k < assylist.size(); k++) {
					setStringCellAndStyle(sheet, "", 7 + k, 5, null, Cell.CELL_TYPE_STRING);// Ａｓｓｙ部番前缀
					setStringCellAndStyle(sheet, "", 7 + k, 27, null, Cell.CELL_TYPE_STRING);// Ａｓｓｙ名称
				}
				// 清空Part List内容
				for (int j = 0; j < 61; j++) {
					setStringCellAndStyle(sheet, "", 23 + j, 1, null, Cell.CELL_TYPE_STRING);// 标号
					setStringCellAndStyle(sheet, "", 23 + j, 13, null, Cell.CELL_TYPE_STRING);// 安装顺序
					setStringCellAndStyle(sheet, "", 23 + j, 20, null, Cell.CELL_TYPE_STRING);// 部品番号
					setStringCellAndStyle(sheet, "", 23 + j, 36, null, Cell.CELL_TYPE_STRING);// 部品名称
					setStringCellAndStyle(sheet, "", 23 + j, 83, null, Cell.CELL_TYPE_STRING);// 数量
					setStringCellAndStyle(sheet, "", 23 + j, 59, null, Cell.CELL_TYPE_STRING);// 板厚
					setStringCellAndStyle(sheet, "", 23 + j, 63, null, Cell.CELL_TYPE_STRING);// 材质
				}
			}
			if (gcnum < page) {
				for (int i = 0; i < page - gcnum; i++) {
					XSSFSheet newsheet = book.cloneSheet(sheetAtIndex);
					book.setSheetOrder(newsheet.getSheetName(), index);
					index++;
				}
			}

		} else {
			// 如果page大于1，则需要复制sheet页
			if (page > 1) {
				for (int i = 1; i < page; i++) {
					XSSFSheet newsheet = book.cloneSheet(sheetAtIndex);
					book.setSheetOrder(newsheet.getSheetName(), index);
					index++;
				}
			}
		}
		/***********************************************/
		// 写构成表数据

		int shnum = 0;
		for (int i = sheetAtIndex; i < index; i++) {

			XSSFSheet sheet = book.getSheetAt(i);

			// 写公共部分
			setStringCellAndStyle(sheet, baseinfo[3], 3, 1, null, Cell.CELL_TYPE_STRING);// 车型
			setStringCellAndStyle(sheet, baseinfo[2], 3, 101, null, Cell.CELL_TYPE_STRING);// 编制
			setStringCellAndStyle(sheet, baseinfo[6], 3, 109, null, Cell.CELL_TYPE_STRING);// 日期
			setStringCellAndStyle(sheet, baseinfo[3], 85, 47, null, Cell.CELL_TYPE_STRING);// 车型
			setStringCellAndStyle(sheet, baseinfo[4], 86, 56, null, Cell.CELL_TYPE_STRING);// 生产线

			// 写partlist上部分信息
			for (int k = 0; k < assylist.size(); k++) {
				String[] val = (String[]) assylist.get(k);
				setStringCellAndStyle(sheet, val[0], 7 + k, 5, null, Cell.CELL_TYPE_STRING);// Ａｓｓｙ部番前缀
				setStringCellAndStyle(sheet, val[1], 7 + k, 27, null, Cell.CELL_TYPE_STRING);// Ａｓｓｙ名称
			}
			// 写partlist信息
			if (i == index - 1) {
				for (int j = 0; j + 61 * shnum < partlist.size(); j++) {
					String[] str = (String[]) partlist.get(j + 61 * shnum);
					// 判断是否为空行
					if (str[7] != null) {
						setStringCellAndStyle(sheet, str[7], 23 + j, 1, null, Cell.CELL_TYPE_STRING);// 标号
						setStringCellAndStyle(sheet, str[0], 23 + j, 13, null, Cell.CELL_TYPE_STRING);// 安装顺序
						setStringCellAndStyle(sheet, str[1], 23 + j, 20, null, Cell.CELL_TYPE_STRING);// 部品番号
						setStringCellAndStyle(sheet, str[2], 23 + j, 36, null, Cell.CELL_TYPE_STRING);// 部品名称
						setStringCellAndStyle(sheet, str[3], 23 + j, 83, null, Cell.CELL_TYPE_STRING);// 数量
						setStringCellAndStyle(sheet, str[4], 23 + j, 59, null, Cell.CELL_TYPE_STRING);// 板厚
						setStringCellAndStyle(sheet, str[5], 23 + j, 63, null, Cell.CELL_TYPE_STRING);// 材质
						// setStringCellAndStyle(sheet, str[6], 23 + j, 72, style,
						// Cell.CELL_TYPE_STRING);// 部品来源

					} else {
					}

				}
			} else {
				for (int j = 0; j + 61 * shnum < 61 + 61 * shnum; j++) {
					// 判断是否为空行
					String[] str = (String[]) partlist.get(j + 61 * shnum);
					// 判断是否为空行
					if (str[7] != null) {
						setStringCellAndStyle(sheet, str[7], 23 + j, 1, null, Cell.CELL_TYPE_STRING);// 标号
						setStringCellAndStyle(sheet, str[0], 23 + j, 13, null, Cell.CELL_TYPE_STRING);// 安装顺序
						setStringCellAndStyle(sheet, str[1], 23 + j, 20, null, Cell.CELL_TYPE_STRING);// 部品番号
						setStringCellAndStyle(sheet, str[2], 23 + j, 36, null, Cell.CELL_TYPE_STRING);// 部品名称
						setStringCellAndStyle(sheet, str[3], 23 + j, 83, null, Cell.CELL_TYPE_STRING);// 数量
						setStringCellAndStyle(sheet, str[4], 23 + j, 59, null, Cell.CELL_TYPE_STRING);// 板厚
						setStringCellAndStyle(sheet, str[5], 23 + j, 63, null, Cell.CELL_TYPE_STRING);// 材质
						// setStringCellAndStyle(sheet, str[6], 23 + j, 72, style,
						// Cell.CELL_TYPE_STRING);// 部品来源
					} else {
					}
				}
			}
			shnum++;
		}
	}

	/*
	 * 写入主数据
	 */
	private void writeMainDataToSheet(XSSFWorkbook book, String[] baseinfo) {
		// TODO Auto-generated method stub

		int sheetnums = book.getNumberOfSheets();
		// 先循环移除打印区域
		for (int i = 0; i < sheetnums; i++) {
			book.removePrintArea(i);
		}

		for (int i = 0; i < sheetnums; i++) {
			XSSFSheet sheet = book.getSheetAt(i);
			String sheetname = sheet.getSheetName();
			if (sheetname.contains("P1_表紙")) {

				setStringCellAndStyle(sheet, baseinfo[1], 4, 5, null, Cell.CELL_TYPE_STRING);// 部门
				setStringCellAndStyle(sheet, baseinfo[3], 14, 1, null, Cell.CELL_TYPE_STRING);// 车型
				setStringCellAndStyle(sheet, baseinfo[4], 20, 2, null, Cell.CELL_TYPE_STRING);// 生产线
				setStringCellAndStyle(sheet, baseinfo[11], 24, 3, null, Cell.CELL_TYPE_STRING);// 部品番号
				setStringCellAndStyle(sheet, baseinfo[5], 27, 3, null, Cell.CELL_TYPE_STRING);// 部品名称
				setStringCellAndStyle(sheet, baseinfo[0], 50, 0, null, Cell.CELL_TYPE_STRING);// 科室
				setStringCellAndStyle(sheet, baseinfo[2], 50, 9, null, Cell.CELL_TYPE_STRING);// 操作人
				setStringCellAndStyle(sheet, baseinfo[6], 43, 11, null, Cell.CELL_TYPE_STRING);// 日期
				setStringCellAndStyle(sheet, baseinfo[2], 42, 24, null, Cell.CELL_TYPE_STRING);// 操作人

				// 设置打印区域
				book.setPrintArea(i, 0, 24, 0, 58);
				PrintSetup printSetup = sheet.getPrintSetup();
				printSetup.setPaperSize(XSSFPrintSetup.A4_PAPERSIZE);
				printSetup.setScale((short) 65);// 自定义缩放，此处100为无缩放
				printSetup.setLandscape(true); // 打印方向，true：横向，false：纵向(默认)
			}
			if (sheetname.contains("P2_目次")) {
				setStringCellAndStyle(sheet, baseinfo[4], 4, 9, null, Cell.CELL_TYPE_STRING);// 生产线
				// 设置打印区域
				book.setPrintArea(i, 0, 62, 0, 50);
				PrintSetup printSetup = sheet.getPrintSetup();
				printSetup.setPaperSize(XSSFPrintSetup.A4_PAPERSIZE);
				printSetup.setScale((short) 68);// 自定义缩放，此处100为无缩放
				printSetup.setLandscape(true); // 打印方向，true：横向，false：纵向(默认)
			}
			if (sheetname.contains("P3_基本仕様")) {
				setStringCellAndStyle(sheet, baseinfo[3], 0, 8, null, Cell.CELL_TYPE_STRING);// 车型
				setStringCellAndStyle(sheet, baseinfo[4], 1, 2, null, Cell.CELL_TYPE_STRING);// 生产线
				setStringCellAndStyle(sheet, baseinfo[3], 6, 3, null, Cell.CELL_TYPE_STRING);// 车型
				// setStringCellAndStyle(sheet, baseinfo[7], 7, 3, null,
				// Cell.CELL_TYPE_STRING);// 治具制作数量
				setStringCellAndStyle(sheet, baseinfo[10], 25, 9, null, Cell.CELL_TYPE_STRING);// 枪型号
				setStringCellAndStyle(sheet, baseinfo[9], 36, 6, null, Cell.CELL_TYPE_STRING);// GAUGE板厚
				// 画图的顶级管理器对象HSSFPatriarch, 一个sheet只能获取一个
				XSSFDrawing hssfPatriarch = (XSSFDrawing) sheet.createDrawingPatriarch();
				XSSFClientAnchor anchor1 = null;
				if (baseinfo[8] != null && baseinfo[8] == "1") {
					anchor1 = new XSSFClientAnchor(-250000, 0, -400000, 0, (short) 6, 25, (short) 7, 26);
				}
				if (baseinfo[8] != null && baseinfo[8] == "2") {
					anchor1 = new XSSFClientAnchor(-150000, 0, -250000, 0, (short) 5, 25, (short) 6, 26);
				}
				// 创建一个椭圆
				if (anchor1 != null) {
					anchor1.setAnchorType(2);
					XSSFSimpleShape ellipse = hssfPatriarch.createSimpleShape(anchor1);
					ellipse.setShapeType(ShapeTypes.ELLIPSE);
					ellipse.setLineStyleColor(0, 0, 0);
					ellipse.setLineWidth(0.75);
					ellipse.setNoFill(false);
				}
				// 设置打印区域
				book.setPrintArea(i, 1, 12, 0, 58);
				PrintSetup printSetup = sheet.getPrintSetup();
				printSetup.setPaperSize(XSSFPrintSetup.A4_PAPERSIZE);
				printSetup.setScale((short) 46);// 自定义缩放，此处100为无缩放
				printSetup.setLandscape(true); // 打印方向，true：横向，false：纵向(默认)
			}
			if (sheetname.contains("部品构成图")) {
				setStringCellAndStyle(sheet, baseinfo[3], 46, 37, null, Cell.CELL_TYPE_STRING);// 车型
				setStringCellAndStyle(sheet, baseinfo[4], 47, 37, null, Cell.CELL_TYPE_STRING);// 生产线
				setStringCellAndStyle(sheet, baseinfo[3], 47, 49, null, Cell.CELL_TYPE_STRING);// 车型
				// 设置打印区域
				book.setPrintArea(i, 1, 61, 2, 49);
				PrintSetup printSetup = sheet.getPrintSetup();
				printSetup.setPaperSize(XSSFPrintSetup.A4_PAPERSIZE);
				printSetup.setScale((short) 69);// 自定义缩放，此处100为无缩放
				printSetup.setLandscape(true); // 打印方向，true：横向，false：纵向(默认)
			}
			if (sheetname.contains("构成部品一覧")) {
				// 设置打印区域
				book.setPrintArea(i, 0, 123, 0, 91);
				PrintSetup printSetup = sheet.getPrintSetup();
				printSetup.setPaperSize(XSSFPrintSetup.A4_PAPERSIZE);
				printSetup.setScale((short) 49);// 自定义缩放，此处100为无缩放
				printSetup.setLandscape(true); // 打印方向，true：横向，false：纵向(默认)
			}
			if (sheetname.contains("Gauge layout")) {
				setStringCellAndStyle(sheet, baseinfo[3], 48, 38, null, Cell.CELL_TYPE_STRING);// 车型
				setStringCellAndStyle(sheet, baseinfo[4], 49, 38, null, Cell.CELL_TYPE_STRING);// 生产线
				// 设置打印区域
				book.setPrintArea(i, 1, 63, 2, 51);
				PrintSetup printSetup = sheet.getPrintSetup();
				printSetup.setPaperSize(XSSFPrintSetup.A4_PAPERSIZE);
				printSetup.setScale((short) 66);// 自定义缩放，此处100为无缩放
				printSetup.setLandscape(true); // 打印方向，true：横向，false：纵向(默认)
			}
			if (sheetname.contains("8_Locate List")) {
				// 设置打印区域
				book.setPrintArea(i, 0, 59, 0, 54);
				PrintSetup printSetup = sheet.getPrintSetup();
				printSetup.setPaperSize(XSSFPrintSetup.A4_PAPERSIZE);
				printSetup.setScale((short) 92);// 自定义缩放，此处100为无缩放
				printSetup.setLandscape(true); // 打印方向，true：横向，false：纵向(默认)
			}
			if (sheetname.contains("断面定位形状仕様")) {
				setStringCellAndStyle(sheet, baseinfo[3], 56, 48, null, Cell.CELL_TYPE_STRING);// 车型
				setStringCellAndStyle(sheet, baseinfo[4], 57, 48, null, Cell.CELL_TYPE_STRING);// 生产线
				// 设置打印区域
				book.setPrintArea(i, 0, 82, 0, 60);
				PrintSetup printSetup = sheet.getPrintSetup();
				printSetup.setPaperSize(XSSFPrintSetup.A4_PAPERSIZE);
				printSetup.setScale((short) 49);// 自定义缩放，此处100为无缩放
				printSetup.setLandscape(true); // 打印方向，true：横向，false：纵向(默认)
			}
			if (sheetname.contains("weld layout")) {
				setStringCellAndStyle(sheet, baseinfo[3], 31, 36, null, Cell.CELL_TYPE_STRING);// 车型
				setStringCellAndStyle(sheet, baseinfo[4], 32, 36, null, Cell.CELL_TYPE_STRING);// 生产线
				setStringCellAndStyle(sheet, baseinfo[3], 32, 48, null, Cell.CELL_TYPE_STRING);// 车型
				// 设置打印区域
				book.setPrintArea(i, 1, 61, 3, 34);
				PrintSetup printSetup = sheet.getPrintSetup();
				printSetup.setPaperSize(XSSFPrintSetup.A4_PAPERSIZE);
				printSetup.setScale((short) 61);// 自定义缩放，此处100为无缩放
				printSetup.setLandscape(true); // 打印方向，true：横向，false：纵向(默认)
			}
			if (sheetname.contains("検知LAYOUT")) {
				setStringCellAndStyle(sheet, baseinfo[3], 31, 36, null, Cell.CELL_TYPE_STRING);// 车型
				setStringCellAndStyle(sheet, baseinfo[4], 32, 36, null, Cell.CELL_TYPE_STRING);// 生产线
				setStringCellAndStyle(sheet, baseinfo[3], 32, 48, null, Cell.CELL_TYPE_STRING);// 车型
				// 设置打印区域
				book.setPrintArea(i, 1, 61, 3, 34);
				PrintSetup printSetup = sheet.getPrintSetup();
				printSetup.setPaperSize(XSSFPrintSetup.A4_PAPERSIZE);
				printSetup.setScale((short) 61);// 自定义缩放，此处100为无缩放
				printSetup.setLandscape(true); // 打印方向，true：横向，false：纵向(默认)
			}
			if (sheetname.contains("検知MATRIX")) {
				setStringCellAndStyle(sheet, baseinfo[12], 3, 10, null, Cell.CELL_TYPE_STRING);// 工位
				setStringCellAndStyle(sheet, baseinfo[3], 34, 38, null, Cell.CELL_TYPE_STRING);// 车型
				setStringCellAndStyle(sheet, baseinfo[4], 35, 38, null, Cell.CELL_TYPE_STRING);// 生产线
				// 设置打印区域
				book.setPrintArea(i, 0, 68, 0, 37);
				PrintSetup printSetup = sheet.getPrintSetup();
				printSetup.setPaperSize(XSSFPrintSetup.A4_PAPERSIZE);
				printSetup.setScale((short) 54);// 自定义缩放，此处100为无缩放
				printSetup.setLandscape(true); // 打印方向，true：横向，false：纵向(默认)
			}
			if (sheetname.contains("14_STD GUN Drawing")) {
				// 设置打印区域
				book.setPrintArea(i, 0, 58, 0, 54);
				PrintSetup printSetup = sheet.getPrintSetup();
				printSetup.setPaperSize(XSSFPrintSetup.A4_PAPERSIZE);
				printSetup.setScale((short) 88);// 自定义缩放，此处100为无缩放
				printSetup.setLandscape(true); // 打印方向，true：横向，false：纵向(默认)
			}
		}

	}

	// 对单元格赋值
	public static void setStringCellAndStyle(XSSFSheet sheet, String value, int rowIndex, int cellIndex,
			XSSFCellStyle Style, int celltype) {

		// 对于整型与字符型的区分 10为整型，11为double型

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
		if (Style != null) {
			cell.setCellStyle(Style);
		}

	}

	/*
	 * 加载空模板
	 */
	private XSSFWorkbook creatXSSFWorkbook(InputStream inputStream, String page2, String gwname, boolean RLflag,
			ArrayList partlist) {
		// TODO Auto-generated method stub
		// multflag,RLflag,
		XSSFWorkbook book = null;
		try {
			book = new XSSFWorkbook(inputStream);
			if (!IsUpdate) {
				int sheetnum = book.getNumberOfSheets();
				ArrayList deletelist = new ArrayList();
				ArrayList copylist = new ArrayList();
				for (int i = 0; i < sheetnum; i++) {
					String sheetname = book.getSheetName(i);
					if (RLflag) {
						if (sheetname.contains("部品构成图")) {
							copylist.add(book.getSheetName(i));
						}
						if (sheetname.contains("Gauge layout")) {
							copylist.add(book.getSheetName(i));
						}
					}
					if (sheetname.contains("weld layout")) {
						copylist.add(book.getSheetName(i));
					}
				}
				// 复制多个相同的sheet
				for (int k = 0; k < copylist.size(); k++) {
					String sheetAllname = (String) copylist.get(k);
					int sheetNums = 1;
//					if (sheetAllname.contains("断面定位形状仕様")) {
//						sheetNums = Integer.parseInt(page1);
//					}
					if (sheetAllname.contains("weld layout")) {
						sheetNums = Integer.parseInt(page2);
					}
					if (sheetAllname.contains("部品构成图") || sheetAllname.contains("Gauge layout")) {
						sheetNums = 2;
					}
					int sheetAt = book.getSheetIndex(sheetAllname);
					int index = sheetAt + 1;
					for (int n = 1; n < sheetNums; n++) {
						XSSFSheet newsheet = book.cloneSheet(sheetAt);
						book.setSheetOrder(newsheet.getSheetName(), index);
						index++;
					}
				}
			}

		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		return book;
	}

	/*
	 * 如果为左右工位，需要把对应的工位部品添加到部品partlist中，根据部品名称进行匹配，如何匹配就添加到下一行标号和安装顺序为空
	 */
	private List getRLHStateData(List sortList, List lHlist) {
		// TODO Auto-generated method stub
		if (sortList != null && sortList.size() > 0) {
			for (int i = 0; i < sortList.size(); i++) {
				if (lHlist != null && lHlist.size() > 0) {
					String[] values = (String[]) sortList.get(i);
					String partName = values[2];
					tempPartlist.add(values);
					for (Iterator<String[]> it = lHlist.iterator(); it.hasNext();) {
						String[] vals = it.next();
						String partName2 = vals[2];
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
				} else {
					String[] values = (String[]) sortList.get(i);
					tempPartlist.add(values);
				}
			}
		}
		return lHlist;
	}

	/*
	 * 获取部品信息
	 */
	private List getPartsinformation(TCComponentBOMLine gwbl) throws TCException, AccessException {
		// TODO Auto-generated method stub
		ArrayList install = new ArrayList();
		ArrayList templist = new ArrayList();// 非内制部品
		// 先获取工位下的安装工序下的零件
		install = Util.getChildrenByBOMLine(gwbl, "B8_BIWOperationRevision");

		for (int i = 0; i < install.size(); i++) {
			// 通过首选项获取部品来源
			Map<String, String> partsource = getSizeRule();

			TCComponentBOMLine bl = (TCComponentBOMLine) install.get(i);
			ArrayList bflist = new ArrayList();
			bflist = Util.getChildrenByBOMLine(bl, "DFL9SolItmPartRevision");
			for (int j = 0; j < bflist.size(); j++) {
				String[] info = new String[8];
				TCComponentBOMLine bfbl = (TCComponentBOMLine) bflist.get(j);
				info[0] = Util.getProperty(bfbl, "bl_sequence_no");// 安装顺序
				if (info[0].isEmpty()) {
					info[0] = "0";
				}
				info[1] = Util.getProperty(bfbl, "bl_DFL9SolItmPartRevision_dfl9_part_no");// 部品番号
				// info[2] = Util.getProperty(bfbl, "bl_rev_object_name");// 部品名称
				info[2] = Util.getProperty(bfbl.getItemRevision(), "dfl9_CADObjectName");// 部品名称
				info[3] = Util.getProperty(bfbl, "bl_quantity");// 数量
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
				// partresoles = Util.getProperty(bfbl, "B8_NoteManualMark");// 部品来源 待确认
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
					// partresoles = Util.getProperty(bfbl, "B8_NoteIsBiwTrUnit");// 部品来源 待确认
				}
				info[6] = partresoles;
				System.out.println(" 部品来源:" + partresValue);
				if (partresValue.equals("Stamping")) {
					String thick = Util.getProperty(bfbl, "bl_DFL9SolItmPartRevision_dfl9PartThickness");// 板厚
					if (Util.isNumber(thick)) {
						Double th = Double.parseDouble(thick);
						info[4] = String.format("%.2f", th);
					} else {
						info[4] = thick;
					}
					info[5] = Util.getProperty(bfbl, "bl_DFL9SolItmPartRevision_dfl9PartMaterial");// 材质
					System.out.println(" 材质:" + Util.getProperty(bfbl, "bl_DFL9SolItmPartRevision_dfl9PartMaterial"));
				} else {
					info[4] = "";// 板厚
					info[5] = "";// 材质
				}
				templist.add(info);
			}
		}

		// 如果零件号相同，合并为一行，数量合计
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
	 * 获取工位的前驱工位的assy部番
	 */
	private List getLastStationPartList(TCComponentBOMLine bl) throws TCException, AccessException {
		List templist = new ArrayList();

		// 内部连接在获取工位的上一个工位的assy部番
		TCProperty pp = bl.getTCProperty("Mfg0predecessors");
		if (pp != null) {
			TCComponent[] obj = pp.getReferenceValueArray();
			for (int i = 0; i < obj.length; i++) {
				TCComponentBOMLine prebl = (TCComponentBOMLine) obj[i];
				String sequence_no = Util.getProperty(prebl, "bl_sequence_no");// 安装顺序
				if (sequence_no.isEmpty()) {
					sequence_no = "0";
				}
				String quantity = Util.getProperty(prebl, "bl_quantity");// 数量
				if (quantity == null || quantity.isEmpty()) {
					quantity = "1";
				}
				// 获取部品信息 ,部品名称为产线名称
				String linename = Util.getProperty(prebl.parent(), "bl_rev_object_name");
				String assyname = linename;

				TCProperty p = prebl.getItemRevision().getTCProperty("b8_ProcAssyNo2");
				String[] assynos;
				if (p != null) {
					assynos = p.getStringValueArray();// Ａｓｓｙ 部番
				} else {
					assynos = null;
				}
				if (assynos != null && assynos.length > 0) {
					for (int j = 0; j < assynos.length; j++) {
						String[] info = new String[8];
						info[0] = sequence_no;// 安装顺序
						info[1] = assynos[j];// 部品番号
						info[2] = assyname;// 部品名称
						info[3] = quantity;// 数量
						info[4] = "";// 板厚
						info[5] = "";// 材质
						info[6] = "内制总成";// 部品来源 待确认
						templist.add(info);
					}
				}
			}
		}
		// 外部连接的工位的上一个工位的assy部番
		List<IMfgFlow> list = FlowUtil.getScopeInputFlows(bl);
		if (list != null && list.size() > 0) {
			for (IMfgFlow flow : list) {
				IMfgNode node = flow.getPredecessor();
				TCComponentBOMLine preComp = (TCComponentBOMLine) node.getComponent();
				String sequence_no = Util.getProperty(preComp, "bl_sequence_no");// 安装顺序
				if (sequence_no.isEmpty()) {
					sequence_no = "0";
				}
				String quantity = Util.getProperty(preComp, "bl_quantity");// 数量
				if (quantity == null || quantity.isEmpty()) {
					quantity = "1";
				}
				// 获取部品信息 ,部品名称为产线名称
				String linename = Util.getProperty(preComp.parent(), "bl_rev_object_name");
				String assyname = linename;

				TCProperty p = preComp.getItemRevision().getTCProperty("b8_ProcAssyNo2");
				String[] assynos;
				if (p != null) {
					assynos = p.getStringValueArray();// Ａｓｓｙ 部番
				} else {
					assynos = null;
				}
				if (assynos != null && assynos.length > 0) {
					for (int j = 0; j < assynos.length; j++) {
						String[] info = new String[8];
						info[0] = sequence_no;// 安装顺序
						info[1] = assynos[j];// 部品番号
						info[2] = assyname;// 部品名称
						info[3] = quantity;// 数量
						info[4] = "";// 板厚
						info[5] = "";// 材质
						info[6] = "内制总成";// 部品来源 待确认
						templist.add(info);
					}
				}
			}
		}

		return templist;
	}

	/*
	 * 设置标号并排序
	 */
	private ArrayList SetLabelsAndSort(List list, TCComponentBOMLine gwbl, TCComponentBOMLine ssgwbl, List lHlist)
			throws AccessException, TCException {

		// 获取完后，对数据进行排序处理
		ArrayList oneList = new ArrayList();
		if (list == null) {
			return null;
		}

		Comparator comparator = getComParatorBysequenceno();
		Collections.sort(list, comparator);

		int label = 0; // 标号标记
		int num = 1;// 标记同种标号的数据行数
		int Occupynum = 0;// 安装顺序为0的占用标号的顺序
		String prePartno = "";// 部品番号前5位标记
		String[] bh = { "A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S",
				"T", "U", "V", "W", "X", "Y", "Z", "AA", "AB", "AC", "AD", "AE", "AF", "AG", "AH", "AI", "AJ", "AK",
				"AL", "AM", "AN", "AO", "AP", "AQ", "AR", "AS", "AT", "AU", "AV", "AW", "AX", "AY", "AZ" };
		// 标号处理
		Map<String, String> tempmap = new HashMap<String, String>();
		List tempPartlist1 = new ArrayList();
		for (int i = 0; i < list.size(); i++) {
			String[] str = (String[]) list.get(i);
			if (str[1].toString().length() > 5) {
				prePartno = str[1].toString().substring(0, 5);
			} else {
				prePartno = str[1].toString();
			}
			String note = tempmap.get(prePartno);
			// 部品番号前5位一样，则标号相同
			if (note != null && !note.isEmpty()) {
				str[7] = note;
				int spno = 0;
				for (int j = 0; j < bh.length; j++) {
					if (bh[j].equals(note)) {
						spno = j + 1 - Occupynum;
					}
				}
				if (!str[0].equals("0")) {
					str[0] = Integer.toString(spno); // 安装顺序重新定义
				}
				String strnum = fymap.get(note);
				int newnum = Integer.parseInt(strnum) + 1;
				fymap.put(note, Integer.toString(newnum));
			} else {
				if (label < 52) {
					str[7] = bh[label];
					if (!str[0].equals("0")) {
						str[0] = Integer.toString(label + 1 - Occupynum); // 安装顺序重新定义
					} else {
						Occupynum++;
					}
				} else {
					str[7] = "";
					System.out.println("超过了规定的标号。。。。");
				}
				fymap.put(bh[label], "1");
				tempmap.put(prePartno, bh[label]);
				label++;
			}
			tempPartlist1.add(str);

		}

		// 如果为左右工位，需要把对应的工位部品添加到部品tempPartlist中，返回对应工位有特有的部品
		List remainList = getRLHStateData(tempPartlist1, lHlist);

		if (remainList != null && remainList.size() > 0) {
			for (int i = 0; i < remainList.size(); i++) {
				String[] str = (String[]) remainList.get(i);
				if (str[1].toString().length() > 5) {
					prePartno = str[1].toString().substring(0, 5);
				} else {
					prePartno = str[1].toString();
				}
				String note = tempmap.get(prePartno);
				// 部品番号前5位一样，则标号相同
				if (note != null && !note.isEmpty()) {
					str[7] = note;
					int spno = 0;
					for (int j = 0; j < bh.length; j++) {
						if (bh[j].equals(note)) {
							spno = j + 1 - Occupynum;
						}
					}
					if (!str[0].equals("0")) {
						str[0] = Integer.toString(spno); // 安装顺序重新定义
					}
					String strnum = fymap.get(note);
					int newnum = Integer.parseInt(strnum) + 1;
					fymap.put(note, Integer.toString(newnum));
				} else {
					if (label < 52) {
						str[7] = bh[label];
						if (!str[0].equals("0")) {
							str[0] = Integer.toString(label + 1 - Occupynum); // 安装顺序重新定义
						} else {
							Occupynum++;
						}
					} else {
						str[7] = "";
						System.out.println("超过了规定的标号。。。。");
					}
					fymap.put(bh[label], "1");
					tempmap.put(prePartno, bh[label]);
					label++;
				}
				tempPartlist.add(str);
			}
		}

		// 把内制部品放到最后
		List LHlist = getLastStationPartList(gwbl);

		if (LHlist != null && LHlist.size() > 0) {
			for (int i = 0; i < LHlist.size(); i++) {
				String[] strVal = (String[]) LHlist.get(i);
				strVal[7] = bh[label];
				strVal[0] = Integer.toString(label + 1 - Occupynum); // 安装顺序重新定义
				tempPartlist.add(strVal);
			}
		}

		if (ssgwbl != null) {
			List RHlist = getLastStationPartList(ssgwbl);
			if (RHlist != null && RHlist.size() > 0) {
				for (int i = 0; i < RHlist.size(); i++) {
					String[] strVal = (String[]) RHlist.get(i);
					strVal[7] = bh[label];
					strVal[0] = Integer.toString(label + 1 - Occupynum); // 安装顺序重新定义
					tempPartlist.add(strVal);
				}
			}
		}

		// 根据标号排序
		Comparator comparator2 = getComParatorBybh();
		Collections.sort(tempPartlist, comparator2);

		String firstNo = "";
		for (int i = 0; i < tempPartlist.size(); i++) {
			String[] value = (String[]) tempPartlist.get(i);
			if (i == 0) {
				firstNo = value[7];
				oneList.add(value);
			} else {
				if (!firstNo.equals(value[7].toString())) {
					String[] str = new String[8];
					oneList.add(str);
					oneList.add(value);
					firstNo = value[7];
				} else {
					oneList.add(value);
				}
			}
		}
		System.out.println(oneList);
		return oneList;
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
	 * GUN形式、GAUGE板厚
	 */
	private String[] getGunInfomation(TCComponentBOMLine gwbl) {
		// TODO Auto-generated method stub
		String[] str = new String[3];
		boolean flag = false;
		ArrayList list = Util.getChildrenByBOMLine(gwbl, "B8_BIWDiscreteOPRevision");
		if (list != null && list.size() > 0) {
			for (int i = 0; i < list.size(); i++) {
				TCComponentBOMLine chilbl = (TCComponentBOMLine) list.get(i);
				String direaname = Util.getProperty(chilbl, "bl_rev_object_name");
				if (direaname.substring(0, 1).equals("R")) {
					flag = true;
				}
				if (i == 0) {
					str[2] = direaname;
				} else {
					str[2] = str[2] + " " + direaname;
				}
			}
			if (flag) {
				str[0] = "1";// 1代表为自动RSW
				str[1] = "12mm";
			} else {
				str[0] = "2";// 2代表为人工PSW
				str[1] = "19mm";
				String typename = Util.getObjectDisplayName(session, "B8_BIWGun");
				String[] propertys = new String[] { "bl_item_object_type", "bl_item_object_type" };
				String[] values = new String[] { typename, typename };
				ArrayList gunlist = Util.searchBOMLine(gwbl, "OR", propertys, "==", values);
				if (gunlist != null && gunlist.size() > 0) {
					for (int i = 0; i < gunlist.size(); i++) {
						TCComponentBOMLine chilbl = (TCComponentBOMLine) gunlist.get(i);
						String gunname = Util.getProperty(chilbl, "bl_B8_BIWGunRevision_b8_Model");
						if (i == 0) {
							str[2] = gunname;
						} else {
							str[2] = str[2] + " " + gunname;
						}
					}
				}
			}
		}

		return str;
	}

	/*
	 * 判断工位是否有对称工位
	 */
	private TCComponentBOMLine getSymmetryState(TCComponentBOMLine linebl, String gwname) throws TCException {
		TCComponentBOMLine ssgwbl = null;
		String ProcLinename = Util.getProperty(linebl, "bl_rev_object_name");
		if (ProcLinename.length() > 1) {
			String rl = ProcLinename.substring(ProcLinename.length() - 2, ProcLinename.length());
			System.out.println("左右工位标识：" + rl);
			if (rl.equals("LH") || rl.equals("RH")) {
				String preLinename = ProcLinename.substring(0, ProcLinename.length() - 2);
				System.out.println("产线名称：" + ProcLinename);
				ArrayList list = Util.getChildrenByBOMLine(linebl.parent(), "B8_BIWMEProcLineRevision");
				for (int i = 0; i < list.size(); i++) {
					TCComponentBOMLine plinebl = (TCComponentBOMLine) list.get(i);
					String plinename = Util.getProperty(plinebl, "bl_rev_object_name");
					System.out.println("虚层下的产线：" + plinename);
					if (!plinename.equals(ProcLinename)) {
						if (plinename.length() > 1
								&& plinename.substring(0, plinename.length() - 2).equals(preLinename)) {
							ArrayList gwlist = Util.getChildrenByBOMLine(plinebl, "B8_BIWMEProcStatRevision");
							for (int j = 0; j < gwlist.size(); j++) {
								TCComponentBOMLine bl = (TCComponentBOMLine) gwlist.get(j);
								String statename = Util.getProperty(bl, "bl_rev_object_name");
								// 如果工位名称中也有左右，也需要区分左右匹配，否则直接按照名称相同匹配
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

	// 生成的报表
	public void saveFiles(String datasetname, String filename, TCComponentBOMLine topbomline,
			TCComponentBOMLine topbomline2, TCComponentItemRevision oldrev) {
		try {
			TCComponentItemRevision toprev = topbomline.getItemRevision();

			String fullFileName = FileUtil.getReportFileName(filename);
			TCComponentDataset ds = Util.createDataset(session, datasetname, fullFileName, "MSExcelX", "excel");

			if (oldrev != null) {
				//判断文档版本是否已发布，如果发布了，先自动升版
				if(oldrev.getDateProperty("date_released") != null) {
					DeepCopyInfo deepCopyInfo = new DeepCopyInfo(oldrev, 1, null, null, false, false, false);
					deepCopyInfo.setAction(2);
					
					TCComponentItemRevision newRev = oldrev.saveAs("", oldrev.getStringProperty("object_name"),  oldrev.getStringProperty("object_desc"), false, new DeepCopyInfo[]{deepCopyInfo});
					
					oldrev = newRev;
				}							
				// 先清空文档下的数据集
				// 移除的时候，需要将所有符合条件的都查找出来，再移除
				TCComponent[] children = TCComponentUtils.getCompsByRelation(oldrev, "IMAN_specification");
				for (TCComponent child : children) {
					if (child instanceof TCComponentDataset) {
						TCComponentDataset dataset = (TCComponentDataset) child;
						oldrev.cutOperation("IMAN_specification", new TCComponent[] { dataset });
						try {
							dataset.delete();
						} catch (Exception e2) {

						}
					}
				}
				// 添加文档与数据集的关系
				oldrev.add("IMAN_specification", ds);
				oldrev.lock();
				oldrev.save();
				oldrev.unlock();
				// 将文档指派项目
				Util.assignProjectComp(oldrev, projects);

			} else {
				TCComponentItemType tcccomponentitemtype = (TCComponentItemType) session
						.getTypeComponent("B8_BIWProcDoc");
				TCComponentItem tccomponentitem = tcccomponentitemtype.create("", "", "B8_BIWProcDoc", datasetname,
						"desc", null);
				tccomponentitem.setProperty("b8_BIWProcDocType", "AB");
				tccomponentitem.lock();
				tccomponentitem.save();
				tccomponentitem.unlock();
				TCComponentItemRevision rev = tccomponentitem.getLatestItemRevision();
				// 添加文档与数据集的关系
				rev.add("IMAN_specification", ds);
				rev.lock();
				rev.save();
				rev.unlock();

				// 添加焊装工位与文档的关系
				toprev.add("IMAN_reference", tccomponentitem);
				toprev.lock();
				toprev.save();
				toprev.unlock();
				System.out.println("LH：" + Util.getProperty(toprev, "item_id"));
				if (topbomline2 != null) {
					TCComponentItemRevision gwrev = topbomline2.getItemRevision();
					System.out.println("RH：" + Util.getProperty(gwrev, "item_id"));
					gwrev.add("IMAN_reference", tccomponentitem);
					gwrev.lock();
					gwrev.save();
					gwrev.unlock();
				}
				// 将文档指派项目
				Util.assignProjectComp(rev, projects);
			}

			// 删除中间文件
			File file = new File(fullFileName);
			if (file.isFile()) {
				file.delete();
			}

		} catch (Exception e) {
			e.printStackTrace();
		}
	}

	/*
	 * 根据工位获取3D图片
	 */
	private Map<String, File> getAll3DPictures(TCComponentItemRevision blrev, String type) throws TCException {
		Map<String, File> piclist = new HashMap<String, File>();
		TCComponent[] tccdata = blrev.getRelatedComponents("IMAN_3D_snap_shot");
		for (TCComponent tcc : tccdata) {
			String objectname = Util.getProperty(tcc, "object_name");
			if (type.equals("1")) { // 部品构成图 名称以数字开头
				if (Util.isNumber(objectname.substring(0, 1))) {
					File file = downLoadPicture1(tcc, "ThumbnailImage");
					if (file != null) {
						piclist.put(objectname, file);
					}
				}
			} else {// 断面定位形状仕様 名称以P或L开头
				if (objectname.substring(0, 1).equals("P") || objectname.substring(0, 1).equals("L")) {
					File file = downLoadPicture1(tcc, "ThumbnailImage");
					if (file != null) {
						piclist.put(objectname, file);
					}
				}
			}
		}
		return piclist;
	}

	/**
	 * 下载图片数据集到本地
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
			// System.out.println("图片类型不匹配："+type);
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

	private static String parseExcel(Sheet sheet, int rowStart, int cellStrat) {
		String resultDataList = "";
		// 解析sheet

		// 校验sheet是否合法
		if (sheet == null) {
			return null;
		}

		// 解析每一行的数据，构造数据对象
		for (int rowNum = rowStart; rowNum < rowStart + 11; rowNum++) {
			Row row = (Row) sheet.getRow(rowNum);
			if (null == row) {
				continue;
			}
			String tempstr = convertRowToData(row, cellStrat);
			if (tempstr != null && !tempstr.isEmpty()) {
				if (resultDataList.isEmpty()) {
					resultDataList = tempstr;
				} else {
					resultDataList = resultDataList + " " + tempstr;
				}

			}

		}
		return resultDataList;
	}

	private static String convertRowToData(Row row, int cellStrat) {
		String resultData = "";
		Cell cell;
		// 板件1
		cell = row.getCell(cellStrat);
		String value = convertCellValueToString(cell);
		if (value != null) {
			resultData = value;
		}

		return resultData;
	}

	/**
	 * 将单元格内容转换为字符串
	 * 
	 * @param cell
	 * @return
	 */
	private static String convertCellValueToString(Cell cell) {
		if (cell == null) {
			return null;
		}
		String returnValue = null;
		switch (cell.getCellType()) {
		case Cell.CELL_TYPE_NUMERIC: // 数字
			Double doubleValue = cell.getNumericCellValue();
			// 格式化科学计数法，取一位整数
			DecimalFormat df = new DecimalFormat("0.00");
			returnValue = df.format(doubleValue);
			break;
		case Cell.CELL_TYPE_STRING: // 字符串
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

	// 查询部品类型首选项，获取部品类型信息
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
}
