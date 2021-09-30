package com.dfl.report.splitexcel;

import java.util.ArrayList;
import java.util.Collections;
import java.util.Comparator;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import com.dfl.report.ExcelReader.WeldPointBoardInformation;
import com.dfl.report.util.Util;
import com.teamcenter.rac.aif.kernel.AIFComponentContext;
import com.teamcenter.rac.aif.kernel.InterfaceAIFComponent;
import com.teamcenter.rac.kernel.TCComponent;
import com.teamcenter.rac.kernel.TCComponentFolder;
import com.teamcenter.rac.kernel.TCComponentItem;
import com.teamcenter.rac.kernel.TCComponentItemRevision;
import com.teamcenter.rac.kernel.TCException;
import com.teamcenter.rac.kernel.TCSession;

public class MoveSplitFilesOperation {

	private TCSession session;
	private List<TCComponentFolder> splitlist;
	private TCComponent savefolder;
	private String resultmessage;
	private Map<TCComponent,TCComponent> splitmap; //需要移动的文档集合
	private Map<TCComponent,TCComponent> map; //存放文件夹下的文档集合
	private Map<String,TCComponent> namemap; //存放文件夹下的文档集合
	private Map<String,TCComponent> foldermap; //存放文件夹下的文件夹集合
	
	public MoveSplitFilesOperation(TCSession session, List<TCComponentFolder> splitlist, TCComponent savefolder) {
		// TODO Auto-generated constructor stub
		this.session = session;
		this.splitlist = splitlist;
		this.savefolder = savefolder;
	}

	public void executeOperation() throws TCException {
		// TODO Auto-generated method stub
		//遍历需要移动的文件的文件夹，获取其下的拆分文档
		TCComponentItemRevision rev;
		splitmap = new HashMap<TCComponent,TCComponent>();
		map = new HashMap<TCComponent,TCComponent>();
		namemap = new HashMap<String,TCComponent>();
		foldermap = new HashMap<String,TCComponent>();
		{
			try {
				//先获取存放文件夹下的文档对象
				getParentDFLdocument(savefolder);
				System.out.println(map);
				for(int i=0;i<splitlist.size();i++) {
					TCComponent spiltfolder = splitlist.get(i);
					//递归遍历获取所有的工艺文档版本对象
					getDFLdocumentobject(spiltfolder);
					AIFComponentContext[] parentfolders = spiltfolder.whereReferenced();
					System.out.println(parentfolders);
					if(parentfolders!=null && parentfolders.length>0) {
						for(AIFComponentContext aif: parentfolders) {
							TCComponent tccfold = (TCComponent) aif.getComponent();
							tccfold.cutOperation("contents", new TCComponent[]{spiltfolder});
						}
					}
				}
				System.out.println(splitmap);
				//保存需要重新排序的文件夹
				List flist = new ArrayList();
				for(Map.Entry<TCComponent,TCComponent> entry: splitmap.entrySet()) {
					TCComponent key = entry.getKey();
					TCComponent value = entry.getValue();
					String objectname = Util.getProperty(key, "object_name");
					//如果按照名称找到，则执行替换操作
					if(namemap.containsKey(objectname)) {
						TCComponent retcc = namemap.get(objectname);
						TCComponent savetcc = map.get(retcc);
						savetcc.cutOperation("contents", new TCComponent[]{retcc});
						savetcc.add("contents", key);
						if(!flist.contains(savetcc)) {
							flist.add(savetcc);
						}
					}else { //没有找到，将文档直接添加到文件夹名称去掉时间与存放文件夹中相同文件夹名称下
						String folder = Util.getProperty(value, "object_name");
						if(folder.length()>12) {
							String tempname = folder.substring(0, folder.length()-12);
							if(foldermap.containsKey(tempname)) {
								TCComponent foldtcc = foldermap.get(tempname);
								foldtcc.add("contents", key);
								if(!flist.contains(foldtcc)) {
									flist.add(foldtcc);
								}
							}
							else {
								savefolder.add("contents", key);
								if(!flist.contains(savefolder)) {
									flist.add(savefolder);
								}
							}
						}else {
							savefolder.add("contents", key);
							if(!flist.contains(savefolder)) {
								flist.add(savefolder);
							}
						}
					}
				}
				//需要根据文档名称排序
				for(int j=0;j<flist.size();j++) {
					TCComponent value = (TCComponent) flist.get(j);
					List sortlist = new ArrayList();
					//先获取文件夹下所有文档版本对象,然后先剪切，排好序在添加到文件夹下
					AIFComponentContext[] contexts = value.getRelated("contents");
					for(AIFComponentContext aif: contexts) {
						TCComponent tcc = (TCComponent) aif.getComponent();
						String objecttype = tcc.getType();
						//如果是工艺文档，就添加到集合，如果是文件夹，继续往下遍历，其他情况不处理
						if(objecttype.equals("DFL9MEDocumentRevision")) {
							Object[] obj = new Object[2];
							String name = Util.getProperty(tcc, "object_name");
							obj[0] = name;
							obj[1] = tcc;
							sortlist.add(obj);						
							value.cutOperation("contents", new TCComponent[]{tcc});
						}
					}
					//先按照sheet名称除序号的部分排序，按照倒叙
//					Comparator comparator1 = getComParatorBypartname();
//					Collections.sort(sortlist, comparator1);
					
					Comparator comparator = getComParatorByname();
					Collections.sort(sortlist, comparator);
					for(int i=0;i<sortlist.size();i++) {
						Object[] objValue = (Object[]) sortlist.get(i);
						TCComponent foldvalue = (TCComponent) objValue[1];
						value.add("contents", foldvalue);
					}	
				}
			}catch(Exception e) {
				resultmessage = e.toString();
			}
			
		}
		return;
	}

	private void getDFLdocumentobject(TCComponent folder) throws TCException {
		// TODO Auto-generated method stub
		AIFComponentContext[] childs = folder.getRelated("contents");
		for(AIFComponentContext aif: childs) {
			TCComponent tcc = (TCComponent) aif.getComponent();
			String objecttype = tcc.getType();
			//如果是工艺文档，就添加到集合，如果是文件夹，继续往下遍历，其他情况不处理
			if(objecttype.equals("DFL9MEDocumentRevision")) {
				splitmap.put(tcc, folder);
			}else if(objecttype.equals("Folder")) {
				getDFLdocumentobject(tcc);
			}else {
				
			}		
		}
	}
	private void getParentDFLdocument(TCComponent folder) throws TCException {
		// TODO Auto-generated method stub
		AIFComponentContext[] childs = folder.getRelated("contents");
		for(AIFComponentContext aif: childs) {
			TCComponent tcc = (TCComponent) aif.getComponent();
			String objecttype = tcc.getType();
			//如果是工艺文档，就添加到集合，如果是文件夹，继续往下遍历，其他情况不处理
			if(objecttype.equals("DFL9MEDocumentRevision")) {
				map.put(tcc, folder);
				String name = Util.getProperty(tcc, "object_name");
				namemap.put(name,tcc);
				String foldername = Util.getProperty(folder, "object_name");
				foldermap.put(foldername, folder);
			}else if(objecttype.equals("Folder")) {
				getParentDFLdocument(tcc);
			}else {
				
			}		
		}
	}

	private Comparator getComParatorByname() {
		// TODO Auto-generated method stub
		Comparator comparator = new Comparator() {

			public int compare(Object obj, Object obj1) {
				// System.setProperty("java.util.Arrays.useLegacyMergeSort", "true");
				Object[] comp1 =  (Object[]) obj;
				Object[] comp2 =  (Object[]) obj1;

				String d1 = "";
				String d2 = "";
				if (obj != null && comp1[0] != null && !comp1[0].toString().isEmpty()) {
//					String[] str = comp1[0].toString().split("-");				
//					d1 = getTmplateSheetname(str).substring(0, 2);
					d1 = comp1[0].toString();
				}
				if (obj1 != null && comp2[0] != null && !comp2[0].toString().isEmpty()) {
//					String[] str = comp2[0].toString().split("-");				
//					d2 = getTmplateSheetname(str).substring(0, 2);
					d2 = comp2[0].toString();
				}
				return d1.compareTo(d2);
			}
		};

		return comparator;
	}
	private Comparator getComParatorBypartname() {
		// TODO Auto-generated method stub
		Comparator comparator = new Comparator() {

			public int compare(Object obj, Object obj1) {
				// System.setProperty("java.util.Arrays.useLegacyMergeSort", "true");
				Object[] comp1 =  (Object[]) obj;
				Object[] comp2 =  (Object[]) obj1;

				String d1 = "";
				String d2 = "";
				if (obj != null && comp1[0] != null && !comp1[0].toString().isEmpty()) {
					String[] str = comp1[0].toString().split("-");				
					d1 = getIsABCEF(getTmplateSheetname(str).substring(2, 3));
				}
				if (obj1 != null && comp2[0] != null && !comp2[0].toString().isEmpty()) {
					String[] str = comp2[0].toString().split("-");				
					d2 = getIsABCEF(getTmplateSheetname(str).substring(2, 3));
				}
				return d1.compareTo(d2);
			}
		};

		return comparator;
	}
	private String getIsABCEF(String str) {
		String value = "";
		char c = str.charAt(0);
		if (c >= 'A' && c <= 'G') {
			value = str;
		}else {
			value = "";
		}
		return value;
	}
	//判断按照“-”截取的sheet是不是点焊，如果是点焊取倒数第二段字符串
	private String getTmplateSheetname(String[] str) {
		String value = "";
		if("MSW".equals(str[str.length-1]) || "PSW".equals(str[str.length-1]) || "RSW".equals(str[str.length-1]) || "SSW".equals(str[str.length-1])) {
			value = str[str.length-2];
		}else {
			value = str[str.length-1];
		}
		return  value;
	}
	
	public String getResultMessage() {
		// TODO Auto-generated method stub
		return resultmessage;
	}

}
