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
	SimpleDateFormat df2 = new SimpleDateFormat("yyyyMMdd HH");// �������ڸ�ʽ

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
		
		// ����������ҵ����󣬻�ȡ���°汾
		// ��ȡ������ҵ���Excel���ݼ�
		{
			TCComponentFolderType foldertype = (TCComponentFolderType) session.getTypeComponent("Folder");
			TCComponentItem item;
			TCComponentItemRevision rev;
			TCComponent[] datasets;
			TCComponent dataset;
			String type;
			// ������·
			{
				Util.callByPass(session, true);
			}
			//��ȡ��Ҫ�滻�������ַ�
			List speciallist = getSpecialChar();
			
			for (int i = 0; i < documentList.size(); i++) {
				item = documentList.get(i);
				rev = item.getLatestItemRevision();//ȡ���°汾
				dataset = null;
				// ��ȡ�汾�µ����ݼ�
				datasets = Util.getRelComponents(rev, "IMAN_specification");
				if (datasets == null || datasets.length <= 0) {
					// resultmessage = " δ��ȡ��Excel���ݼ�����";
					System.out.println(rev.toDisplayString() + " δ��ȡ��Excel���ݼ�����");
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
					// resultmessage = " δ��ȡ��Excel���ݼ�����";
					System.out.println(rev.toDisplayString() + " δ��ȡ��Excel���ݼ�����");
					continue;
				}
				String tempPath = getTempPath();
				File dirfile = new File(tempPath);
				if (!dirfile.exists()) {
					dirfile.mkdir();
				}
				// �������ݼ��ļ�
				File file = downloadFile((TCComponentDataset) dataset, tempPath);
				if (file == null) {
					// resultmessage = " �������ݼ��ļ�����";
					System.out.println(rev.toDisplayString() + " �������ݼ��ļ�����");
					continue;
				}
				// �ж��ĵ���ʲô���͵ı���
				int doctype = getDocumentType(rev);
				String prefilename = "";
				// �����ĵ�������ֵ����ȡ�ļ�ǰ׺����
				String foldername2 = Util.getProperty(item, "dfl9_vehiclePlant");
				String foldername1 = Util.getProperty(item, "dfl9_processArea");
				String reporttype = "";
				if (doctype == 1) {
					prefilename = foldername2 + "-" + foldername1 + "-" + Util.getProperty(rev, "object_name") + "-";
					reporttype = "AB";
				} else if (doctype == 2) {
					prefilename = "������ͼA��-" + foldername2 + "-";
					reporttype = "GG-A";
				} else if (doctype == 4){	
					
					prefilename = "ֱ���嵥" + getDocName(file) + "-";
					reporttype = "ZC";
				}else
				{
					prefilename = "�˹���ǯ���Ӳ�����-" + foldername2 + "-";
					reporttype = "CS";
				}
				// prefilename="HD1H1-02.FF-";
				// ���ñ���vbs���excel��
				File[] files = callVBSProgram(tempPath, file.getAbsolutePath(), vbsFilePath, "");
				String itemname = Util.getProperty(item, "object_name");
				if (files == null || files.length <= 0) {
					if (resultmessage != null) {
						resultmessage = resultmessage + itemname + "VBS��ֽ��Ϊ�գ�\n";
					} else {
						resultmessage = itemname + "VBS��ֽ��Ϊ�գ�\n";
					}

					System.out.println("VBS��ֽ��Ϊ�գ�");
					continue;
				}
				// ����󣬰��ļ�ɾ��
				if (file != null) {
					file.delete();
				}

				// ��ȡ���ǰ�ĵ�������Ŀ
				List<TCComponentProject> projects = new LinkedList<TCComponentProject>();
				TCComponent[] comps = TCComponentProjectUtils.getTCComponentProjects(rev);
				if (comps != null && comps.length > 0) {
					for (TCComponent comp : comps) {
						if (comp instanceof TCComponentProject) {
							projects.add((TCComponentProject) comp);
						}
					}
				}

				// �����ĵ���Ӧ���ļ���
				String itemnmae = Util.getProperty(item, "object_name");
				// String splitname = itemnmae + df2.format(new Date()) + "ʱ";
				String splitname = itemnmae;
				if (doctype == 4) 
				{
					splitname = "ֱ���嵥���";
				}
				// ���ж��Ƿ��Ѿ�������ҵ�����ļ���
				TCComponentFolder Splitfolder = null;// ���ϼ��ļ���
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

				// ���û�оʹ���
				if (Splitfolder == null) {
					Splitfolder = foldertype.create(splitname, "", "Folder");
					savefolder.add("contents", Splitfolder);
				} else { // ���������µ��ĵ��汾�����Ƴ�
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
				System.out.println("����ĵ���Ӧ���ļ��У�" + Util.getProperty(Splitfolder, "object_name"));
				
				
				//huangbin modify start 2021/7/29 
				processStation = SplitExcelUtil.getProcessStation(rev);
				//HashMap<String,TCComponent> opNameMap = getOpNameMap(processStation);
				//huangbin modify end
				
				for (int j = 0; j < files.length; j++) {

					File chilfile = files[j];
					System.out.println("chilfile:" + chilfile.getAbsolutePath());

					// �ļ�������
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
					
					
					
					// �ϴ���ֺ��Excel���½����ݼ�
					// String tempname = newFile.getName();//�����������\��������
					String tempname = prefilename + oldname;
					String filename = tempname.replace(".xlsx", "");
					String fullFileName = newFile.getPath();
					filename = replaceSpecialchar(speciallist,filename);
					System.out.println("fullFileName: " + fullFileName);
					System.out.println("filename: " + filename);
					TCComponentDataset ds = Util.createDatasetKeepFile(session, filename, fullFileName, "MSExcelX", "excel");
					// �����ĵ����󣬲�����excel���ݼ�
					System.out.println("�������ݼ�:" + Util.getProperty(ds, "object_name"));
					if (ds == null) {
						resultmessage = "�������ݼ�ʧ�ܣ�";
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
						resultmessage = "�����ĵ�����ʧ�ܣ�";
						continue;
					}
					System.out.println("�����ĵ�:" + Util.getProperty(dfl9MEDocumentRev, "object_name"));
					// ���ݼ���ӵ��ĵ��汾��
					// TCComponentUtils.createRelation(dfl9MEDocumentRev, ds, "IMAN_specification");
					dfl9MEDocumentRev.add("IMAN_specification", ds);
					// ���������ĵ�ָ�ɵ���Ŀ
					for (TCComponentProject proj : projects) {
						TCComponentProjectUtils.assignProject(dfl9MEDocumentRev, proj);
					}
					// �ĵ��汾��ӵ��ļ���
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
					
					// �������󣬰��ļ�ɾ��
					if (chilfile != null) {
						chilfile.delete();
					}
				}
			}
			// �ر���·
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
	 * { System.out.println(">>>WARN:��λ�����¼�����Ϊ�գ�������"); return map; }
	 * 
	 * //2����ȡ��λ���յ��¼��������汾B8_BIWDiscreteOPRevision��B8_BIWOperationRevision��
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
	

	
	

	// �ж��ĵ���ʲô���͵ı���
	private int getDocumentType(TCComponentItemRevision rev) {
		// TODO Auto-generated method stub
		int type = 0;
		String process_type = Util.getProperty(rev, "dfl9_process_type");
		String process_file_type = Util.getProperty(rev, "dfl9_process_file_type");
		if (("H".equals(process_type) || "��װ����".equals(process_type))
				&& ("AB".equals(process_file_type) || "������ҵ����AB��".equals(process_file_type))) {
			type = 1;
		} else if (("H".equals(process_type) || "��װ����".equals(process_type))
				&& ("GG-A".equals(process_file_type) || "������ͼA��".equals(process_file_type))) {
			type = 2;
		} else if (("H".equals(process_type) || "��װ����".equals(process_type))
				&& ("ZC".equals(process_file_type) || "ֱ�����Ķ����嵥".equals(process_file_type))) {
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
				System.out.println("vbs����ļ�����");
			}

		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}

		return null;
	}

	/**
	 * ͨ�����ƻ�ȡWordģ���ļ�
	 * 
	 * @param name
	 * @return ���ݼ�����
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
	
	// ��ѯ��Ҫ�滻�������ַ�
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

    //��ȡֱ���嵥���ļ�����
	private String getDocName(File file)
	{
		String name = "";
		try {
			InputStream inputStream = new FileInputStream(file);
			Workbook workbook =baseinfoExcelReader.getWorkbook(inputStream, "xlsx");
			Sheet sheet = workbook.getSheetAt(0);
			// У��sheet�Ƿ�Ϸ�
			if (sheet == null) {
				System.out.println("δ�ҵ�����sheetҳ��");
				return name;
			}
			Row row = sheet.getRow(8);
			if(row!=null)
			{
				Cell cell = row.getCell(3);
				if(cell!=null)
				{
					String value = cell.getStringCellValue().trim();
					System.out.println("��ȡ���ļ���ţ�" + value);
					String[] arrStr = value.split("-");
					if(arrStr.length>1)
					{
						if (arrStr[1].length() > 2) {
							//��ȡ���һ��H��λ��
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
								System.out.println("��ȡ���ļ����벻�淶���ļ�����Ϊ��");
								return "";
							}
							//zz1h1
							String factory = "";
							String comp = arrStr[1].substring(0, index-1);
							if("1".equals(arrStr[1].substring(index-1,index)))
							{
								factory = "һ";
							}
							else if("2".equals(arrStr[1].substring(index-1,index)))
							{
								factory = "��";
							}
							else if("3".equals(arrStr[1].substring(index-1,index)))
							{
								factory = "��";
							}
							else if("4".equals(arrStr[1].substring(index-1,index)))
							{
								factory = "��";
							}
							else if("5".equals(arrStr[1].substring(index-1,index)))
							{
								factory = "��";
							}
							else if("6".equals(arrStr[1].substring(index-1,index)))
							{
								factory = "��";
							}
							else if("7".equals(arrStr[1].substring(index-1,index)))
							{
								factory = "��";
							}
							else if("8".equals(arrStr[1].substring(index-1,index)))
							{
								factory = "��";
							}
							else if("9".equals(arrStr[1].substring(index-1,index)))
							{
								factory = "��";
							}
							else
							{
								comp = arrStr[1].substring(0, index);
							}
							name = "-" + comp + factory + "����" + "NO"
									+ arrStr[1].substring(arrStr[1].length() - 1);
						} else {
							System.out.println("��ȡ���ļ����벻�淶���ļ�����Ϊ��");
						}			
					}
					else
					{
						System.out.println("��ȡ���ļ����벻�淶���ļ�����Ϊ��");
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
