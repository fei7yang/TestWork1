package com.dfl.report.splitexcel;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.InputStream;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.HashMap;
import java.util.LinkedList;
import java.util.List;
import java.util.Map;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import com.dfl.report.ExcelReader.baseinfoExcelReader;
import com.dfl.report.common.TCComponentProjectUtils;
import com.dfl.report.common.TCComponentUtils;
import com.dfl.report.util.ReportUtils;
import com.dfl.report.util.Util;
import com.teamcenter.rac.aif.AbstractAIFOperation;
import com.teamcenter.rac.aif.AbstractAIFUIApplication;
import com.teamcenter.rac.aif.kernel.AIFComponentContext;
import com.teamcenter.rac.aifrcp.AIFUtility;
import com.teamcenter.rac.kernel.TCComponent;
import com.teamcenter.rac.kernel.TCComponentDataset;
import com.teamcenter.rac.kernel.TCComponentFolder;
import com.teamcenter.rac.kernel.TCComponentFolderType;
import com.teamcenter.rac.kernel.TCComponentItem;
import com.teamcenter.rac.kernel.TCComponentItemRevision;
import com.teamcenter.rac.kernel.TCComponentProject;
import com.teamcenter.rac.kernel.TCException;
import com.teamcenter.rac.kernel.TCPreferenceService;
import com.teamcenter.rac.kernel.TCSession;
import com.teamcenter.schemas.soa._2006_03.exceptions.ServiceException;
import com.teamcenter.services.rac.core._2008_06.DataManagement.CreateResponse;

public class SplitExcelOperation extends AbstractAIFOperation {

	private TCSession session;
	private ArrayList<TCComponentItem> documentList;
	private String vbsFilePath;
	private TCComponent savefolder;
	private String resultmessage;
	SimpleDateFormat df2 = new SimpleDateFormat("yyyyMMdd HH");// 设置日期格式

	public SplitExcelOperation(TCSession session, ArrayList<TCComponentItem> documentList, String vbsFilePath,
			TCComponent savefolder) {
		// TODO Auto-generated constructor stub
		this.session = session;
		this.documentList = documentList;
		this.vbsFilePath = vbsFilePath;
		this.savefolder = savefolder;
	}

	@Override
	public void executeOperation() throws Exception {
		// TODO Auto-generated method stub
		processStation=null;
		
		// 遍历工程作业表对象，获取最新版本
		// 获取工程作业表的Excel数据集
		{
			TCComponentFolderType foldertype = (TCComponentFolderType) session.getTypeComponent("Folder");
			TCComponentItem item;
			TCComponentItemRevision rev;
			TCComponent[] datasets;
			TCComponent dataset;
			String type;
			// 开启旁路
			{
				Util.callByPass(session, true);
			}
			//获取需要替换的特殊字符
			List speciallist = getSpecialChar();
			
			for (int i = 0; i < documentList.size(); i++) {
				item = documentList.get(i);
				rev = item.getLatestItemRevision();//取最新版本
				dataset = null;
				// 获取版本下的数据集
				datasets = Util.getRelComponents(rev, "IMAN_specification");
				if (datasets == null || datasets.length <= 0) {
					// resultmessage = " 未获取到Excel数据集对象";
					System.out.println(rev.toDisplayString() + " 未获取到Excel数据集对象");
					continue;
				}

				for (int j = 0; j < datasets.length; j++) {
					type = datasets[j].getType();
					if ("MSExcelX".equals(type)) {
						dataset = datasets[j];
						break;
					}
				}
				if (dataset == null) {
					// resultmessage = " 未获取到Excel数据集对象";
					System.out.println(rev.toDisplayString() + " 未获取到Excel数据集对象");
					continue;
				}
				String tempPath = getTempPath();
				File dirfile = new File(tempPath);
				if (!dirfile.exists()) {
					dirfile.mkdir();
				}
				// 下载数据集文件
				File file = downloadFile((TCComponentDataset) dataset, tempPath);
				if (file == null) {
					// resultmessage = " 下载数据集文件错误";
					System.out.println(rev.toDisplayString() + " 下载数据集文件错误");
					continue;
				}
				// 判断文档是什么类型的报表
				int doctype = getDocumentType(rev);
				String prefilename = "";
				// 根据文档的属性值，获取文件前缀名称
				String foldername2 = Util.getProperty(item, "dfl9_vehiclePlant");
				String foldername1 = Util.getProperty(item, "dfl9_processArea");
				String reporttype = "";
				if (doctype == 1) {
					prefilename = foldername2 + "-" + foldername1 + "-" + Util.getProperty(rev, "object_name") + "-";
					reporttype = "AB";
				} else if (doctype == 2) {
					prefilename = "管理工程图A表-" + foldername2 + "-";
					reporttype = "GG-A";
				} else if (doctype == 4){	
					
					prefilename = "直材清单" + getDocName(file) + "-";
					reporttype = "ZC";
				}else
				{
					prefilename = "人工焊钳焊接参数表-" + foldername2 + "-";
					reporttype = "CS";
				}
				// prefilename="HD1H1-02.FF-";
				// 调用本地vbs拆分excel表
				File[] files = callVBSProgram(tempPath, file.getAbsolutePath(), vbsFilePath, "");
				String itemname = Util.getProperty(item, "object_name");
				if (files == null || files.length <= 0) {
					if (resultmessage != null) {
						resultmessage = resultmessage + itemname + "VBS拆分结果为空！\n";
					} else {
						resultmessage = itemname + "VBS拆分结果为空！\n";
					}

					System.out.println("VBS拆分结果为空！");
					continue;
				}
				// 拆完后，把文件删除
				if (file != null) {
					file.delete();
				}

				// 获取拆分前文档所属项目
				List<TCComponentProject> projects = new LinkedList<TCComponentProject>();
				TCComponent[] comps = TCComponentProjectUtils.getTCComponentProjects(rev);
				if (comps != null && comps.length > 0) {
					for (TCComponent comp : comps) {
						if (comp instanceof TCComponentProject) {
							projects.add((TCComponentProject) comp);
						}
					}
				}

				// 创建文档对应的文件夹
				String itemnmae = Util.getProperty(item, "object_name");
				// String splitname = itemnmae + df2.format(new Date()) + "时";
				String splitname = itemnmae;
				if (doctype == 4) 
				{
					splitname = "直材清单拆分";
				}
				// 先判断是否已经存在作业表拆分文件夹
				TCComponentFolder Splitfolder = null;// 上上级文件夹
				AIFComponentContext[] content3 = savefolder.getRelated("contents");
				if (content3 != null && content3.length > 0) {
					for (AIFComponentContext aif : content3) {
						TCComponent tcc = (TCComponent) aif.getComponent();
						String objectname = Util.getProperty(tcc, "object_name");
						if (tcc.isTypeOf("Folder") && objectname.equals(splitname)) {
							Splitfolder = (TCComponentFolder) tcc;
							break;
						}
					}
				}

				// 如果没有就创建
				if (Splitfolder == null) {
					Splitfolder = foldertype.create(splitname, "", "Folder");
					savefolder.add("contents", Splitfolder);
				} else { // 存在则将其下的文档版本对象移除
					AIFComponentContext[] spcontent = Splitfolder.getRelated("contents");
					if (spcontent != null && spcontent.length > 0) {
						for (AIFComponentContext aif : spcontent) {
							TCComponent tcc = (TCComponent) aif.getComponent();
							if (tcc.isTypeOf("DFL9MEDocumentRevision")) {
								Splitfolder.remove("contents", tcc);
							}
						}
					}
				}
				System.out.println("拆分文档对应的文件夹：" + Util.getProperty(Splitfolder, "object_name"));
				
				
				//huangbin modify start 2021/7/29 
				processStation = SplitExcelUtil.getProcessStation(rev);
				//HashMap<String,TCComponent> opNameMap = getOpNameMap(processStation);
				//huangbin modify end
				
				for (int j = 0; j < files.length; j++) {

					File chilfile = files[j];
					System.out.println("chilfile:" + chilfile.getAbsolutePath());

					// 文件重命名
					String oldname = chilfile.getName();
					String oldpath = chilfile.getPath().replace(oldname, "");
					String oldtemp = prefilename + oldname;
					oldtemp = replaceSpecialchar(speciallist,oldtemp);						
					String newfilename = oldpath + oldtemp;
					System.out.println("newfilename:" + newfilename);
					
					File newFile = new File(newfilename);
					if (chilfile.exists() && chilfile.isFile()) {
						chilfile.renameTo(newFile);
					}
					
					
					
					// 上传拆分后的Excel，新建数据集
					// String tempname = newFile.getName();//如果名称中有\会有问题
					String tempname = prefilename + oldname;
					String filename = tempname.replace(".xlsx", "");
					String fullFileName = newFile.getPath();
					filename = replaceSpecialchar(speciallist,filename);
					System.out.println("fullFileName: " + fullFileName);
					System.out.println("filename: " + filename);
					TCComponentDataset ds = Util.createDatasetKeepFile(session, filename, fullFileName, "MSExcelX", "excel");
					// 创建文档对象，并关联excel数据集
					System.out.println("创建数据集:" + Util.getProperty(ds, "object_name"));
					if (ds == null) {
						resultmessage = "创建数据集失败！";
						continue;
					}
					TCComponentItemRevision dfl9MEDocumentRev = null;
					Map<String, Object> itemMap = new HashMap<String, Object>();
					Map<String, Object> itemRevisionMap = new HashMap<String, Object>();
					Map<String, Object> itemRevMasterFormMap = new HashMap<String, Object>();
					itemMap.put("item_id", ""); //$NON-NLS-1$ //$NON-NLS-2$
					itemMap.put("object_name", filename); //$NON-NLS-1$
					itemMap.put("object_desc", ""); //$NON-NLS-1$
					itemMap.put("object_type", "DFL9MEDocument"); //$NON-NLS-1$
					itemRevisionMap.put("object_type", "DFL9MEDocumentRevision"); //$NON-NLS-1$
					itemRevisionMap.put("object_name", filename); //$NON-NLS-1$
					itemRevisionMap.put("dfl9_process_type", "H"); //$NON-NLS-1$
					itemRevisionMap.put("dfl9_process_file_type", reporttype); //$NON-NLS-1$
					itemRevMasterFormMap.put("object_type", "DFL9MEDocumentRevisionMaster"); //$NON-NLS-1$
					try {
						CreateResponse respose = TCComponentUtils.create(itemMap, itemRevisionMap,
								itemRevMasterFormMap);
						int num = respose.serviceData.sizeOfCreatedObjects();
						if (num > 0) {
							for (int k = 0; k < num; k++) {
								TCComponent comp = respose.serviceData.getCreatedObject(k);
								if (comp instanceof TCComponentItemRevision) {
									dfl9MEDocumentRev = (TCComponentItemRevision) comp;
								}
							}
						}
					} catch (ServiceException e) {
						e.printStackTrace();
					}
					if (dfl9MEDocumentRev == null) {
						resultmessage = "创建文档对象失败！";
						continue;
					}
					System.out.println("创建文档:" + Util.getProperty(dfl9MEDocumentRev, "object_name"));
					// 数据集添加到文档版本下
					// TCComponentUtils.createRelation(dfl9MEDocumentRev, ds, "IMAN_specification");
					dfl9MEDocumentRev.add("IMAN_specification", ds);
					// 将创建的文档指派到项目
					for (TCComponentProject proj : projects) {
						TCComponentProjectUtils.assignProject(dfl9MEDocumentRev, proj);
					}
					// 文档版本添加到文件夹
					Splitfolder.add("contents", dfl9MEDocumentRev);


					//huangbin modify start 2021/7/29 
					{
						SplitExcelUtil.setDocRelation(processStation,rev,dfl9MEDocumentRev,newFile);
					}
					//huangbin modify end
					if (newFile.exists())
					{
						newFile.delete();
					}
					
					// 创建完后后，把文件删除
					if (chilfile != null) {
						chilfile.delete();
					}
				}
			}
			// 关闭旁路
			{
				Util.callByPass(session, false);
			}
		}
		return;
	}
	
	/*
	 * private HashMap<String, TCComponent> getOpNameMap(TCComponentItemRevision
	 * processStation2) throws TCException { // TODO Auto-generated method stub
	 * HashMap<String, TCComponent> map = new HashMap<String, TCComponent>();
	 * if(processStation2!=null) { TCComponent[] childrens =
	 * processStation2.getReferenceListProperty("ps_children"); if(childrens==null )
	 * { System.out.println(">>>WARN:工位工艺下级工序为空，跳过！"); return map; }
	 * 
	 * //2、获取工位工艺的下级工序对象版本B8_BIWDiscreteOPRevision、B8_BIWOperationRevision、
	 * B8_BIWArcWeldOPRevision
	 * 
	 * int opCount = childrens.length ; for (int i = 0; i < opCount; i++) {
	 * TCComponent opRev = childrens[i]; String objectType = opRev.getType(); String
	 * opName = Util.getProperty(opRev, "object_name");
	 * 
	 * 
	 * } } return map; }
	 */

	TCComponentItemRevision processStation;
	

	
	

	// 判断文档是什么类型的报表
	private int getDocumentType(TCComponentItemRevision rev) {
		// TODO Auto-generated method stub
		int type = 0;
		String process_type = Util.getProperty(rev, "dfl9_process_type");
		String process_file_type = Util.getProperty(rev, "dfl9_process_file_type");
		if (("H".equals(process_type) || "焊装工艺".equals(process_type))
				&& ("AB".equals(process_file_type) || "工程作业程序AB表".equals(process_file_type))) {
			type = 1;
		} else if (("H".equals(process_type) || "焊装工艺".equals(process_type))
				&& ("GG-A".equals(process_file_type) || "管理工程图A表".equals(process_file_type))) {
			type = 2;
		} else if (("H".equals(process_type) || "焊装工艺".equals(process_type))
				&& ("ZC".equals(process_file_type) || "直材消耗定额清单".equals(process_file_type))) {
			type = 4;
		}else
		{
			type = 3;
		}
		return type;
	}

	private File[] callVBSProgram(String tempPath, String xlsFilePath, String vbsFilePath, String prefilename) {
		// TODO Auto-generated method stub
		String oupFilePath = tempPath + "output";
		File dirfile = new File(oupFilePath);
		if (!dirfile.exists()) {
			dirfile.mkdir();
		}
		final String command = "wscript  \"" + vbsFilePath + "\" \"" + xlsFilePath + "\" " + oupFilePath + " \""
				+ prefilename + "\"";
		System.out.println(command);
		try {
			Process process = Runtime.getRuntime().exec(command);
			try {
				process.waitFor();
			} catch (InterruptedException e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			}
			System.out.println("finish");

			String oupFilePath2 = tempPath + "outputsuccess";

			File file = new File(oupFilePath2);
			if (file.exists()) {
				File[] files = file.listFiles();
//					for (int i = 0; i < files.length; i++) {
//						System.out.println("files:"+files[i].getPath());
//					}
				if (files != null) {
					return files;
				}
			} else {
				System.out.println("vbs拆分文件错误！");
			}

		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}

		return null;
	}

	/**
	 * 通过名称获取Word模板文件
	 * 
	 * @param name
	 * @return 数据集对象
	 */
	public File downloadFile(TCComponentDataset dataset, String difPath) {

		try {
			System.out.println(dataset.getType());
			if (dataset.getType().equals("MSExcelX")) {
				File files[] = dataset.getFiles("excel", difPath);
				if (files == null || files.length <= 0) {
					return null;
				}
				System.err.println(files[0].getPath());
				return files[0];
			}
		} catch (Exception e) {
			e.printStackTrace();
		}
		return null;
	}

	public String getTempPath() {
		String path = "";
		String tmpPath = System.getProperty("java.io.tmpdir");
		// System.out.println("tmpPath:"+tmpPath);
		if (tmpPath.endsWith("\\")) {
			path = tmpPath + new Date().getTime();
		} else {
			path = tmpPath + "\\" + new Date().getTime();
		}
		path = path + "\\";
		System.out.println("tempPath=" + path);
		return path;
	}

	public String getResultMessage() {
		return resultmessage;
	}

	private String replaceSpecialchar(List list,String name)
	{
		if(list!=null && list.size()>0)
		{
			for(int i=0;i<list.size();i++)
			{
				String ch = (String)list.get(i);				
				name = name.replace(ch, "");
			}
		}
		
		return  name;
	}
	
	// 查询需要替换的特殊字符
	private List getSpecialChar() {
		List rule = new ArrayList();
		try {
			AbstractAIFUIApplication app = AIFUtility.getCurrentApplication();
			TCSession session = (TCSession) app.getSession();
			TCPreferenceService preferenceService = session.getPreferenceService();

			String str = preferenceService.getPreferenceDescription("DFL_EngineeringWorkListSplitSheetName");
			if (str != null) {
				String[] values = preferenceService.getStringValues("DFL_EngineeringWorkListSplitSheetName");
				if (values != null) {
					for (int i = 0; i < values.length; i++) {
						String value = values[i];
						value = value.replace("<", "").replace(">", "");
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

    //获取直材清单的文件名称
	private String getDocName(File file)
	{
		String name = "";
		try {
			InputStream inputStream = new FileInputStream(file);
			Workbook workbook =baseinfoExcelReader.getWorkbook(inputStream, "xlsx");
			Sheet sheet = workbook.getSheetAt(0);
			// 校验sheet是否合法
			if (sheet == null) {
				System.out.println("未找到封面sheet页！");
				return name;
			}
			Row row = sheet.getRow(8);
			if(row!=null)
			{
				Cell cell = row.getCell(3);
				if(cell!=null)
				{
					String value = cell.getStringCellValue().trim();
					System.out.println("获取的文件编号：" + value);
					String[] arrStr = value.split("-");
					if(arrStr.length>1)
					{
						if (arrStr[1].length() > 2) {
							//获取最后一个H的位置
							int index = -1;
							for(int i=arrStr[1].length()-1;i>=0;i--)
							{
								char ch = arrStr[1].charAt(i);
								if('H' == ch)
								{
									index = i;
									break;
								}							
							}
							if(index == -1)
							{
								System.out.println("获取的文件编码不规范，文件名称为空");
								return "";
							}
							//zz1h1
							String factory = "";
							String comp = arrStr[1].substring(0, index-1);
							if("1".equals(arrStr[1].substring(index-1,index)))
							{
								factory = "一";
							}
							else if("2".equals(arrStr[1].substring(index-1,index)))
							{
								factory = "二";
							}
							else if("3".equals(arrStr[1].substring(index-1,index)))
							{
								factory = "三";
							}
							else if("4".equals(arrStr[1].substring(index-1,index)))
							{
								factory = "四";
							}
							else if("5".equals(arrStr[1].substring(index-1,index)))
							{
								factory = "五";
							}
							else if("6".equals(arrStr[1].substring(index-1,index)))
							{
								factory = "六";
							}
							else if("7".equals(arrStr[1].substring(index-1,index)))
							{
								factory = "七";
							}
							else if("8".equals(arrStr[1].substring(index-1,index)))
							{
								factory = "八";
							}
							else if("9".equals(arrStr[1].substring(index-1,index)))
							{
								factory = "九";
							}
							else
							{
								comp = arrStr[1].substring(0, index);
							}
							name = "-" + comp + factory + "工厂" + "NO"
									+ arrStr[1].substring(arrStr[1].length() - 1);
						} else {
							System.out.println("获取的文件编码不规范，文件名称为空");
						}			
					}
					else
					{
						System.out.println("获取的文件编码不规范，文件名称为空");
					}
				}
			}
			
		} catch (FileNotFoundException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
				
		return name;
	}
}
