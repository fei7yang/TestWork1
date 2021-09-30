package com.dfl.report.mfcadd;

import java.awt.Toolkit;
import java.awt.datatransfer.Clipboard;
import java.awt.datatransfer.DataFlavor;
import java.awt.datatransfer.StringSelection;
import java.awt.datatransfer.Transferable;
import java.io.BufferedReader;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.InputStreamReader;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.regex.Matcher;
import java.util.regex.Pattern;
import java.util.regex.PatternSyntaxException;

import javax.swing.JTable;

import org.eclipse.jface.dialogs.MessageDialog;
import org.eclipse.swt.graphics.Rectangle;
import org.eclipse.swt.widgets.Display;
import org.eclipse.swt.widgets.Shell;

import com.teamcenter.rac.aifrcp.AIFUtility;
import com.teamcenter.rac.kernel.TCAccessControlService;
import com.teamcenter.rac.kernel.TCComponent;
import com.teamcenter.rac.kernel.TCComponentBOMLine;
import com.teamcenter.rac.kernel.TCComponentDataset;
import com.teamcenter.rac.kernel.TCComponentDatasetType;
import com.teamcenter.rac.kernel.TCComponentFolder;
import com.teamcenter.rac.kernel.TCComponentItemRevision;
import com.teamcenter.rac.kernel.TCComponentListOfValues;
import com.teamcenter.rac.kernel.TCComponentListOfValuesType;
import com.teamcenter.rac.kernel.TCComponentQuery;
import com.teamcenter.rac.kernel.TCComponentQueryType;
import com.teamcenter.rac.kernel.TCComponentTask;
import com.teamcenter.rac.kernel.TCException;
import com.teamcenter.rac.kernel.TCSession;
import com.teamcenter.rac.kernel.TCUserService;
import com.teamcenter.rac.kernel.VariantClause;
import com.teamcenter.rac.kernel.VariantCondition;
import com.teamcenter.rac.kernel.VariantOption;
import com.teamcenter.services.rac.core.DataManagementService;
import com.teamcenter.services.rac.core._2008_06.DataManagement;
import com.teamcenter.services.internal.rac.structuremanagement.VariantManagementService;
import com.teamcenter.services.internal.rac.structuremanagement._2011_06.VariantManagement;
import com.teamcenter.services.internal.rac.structuremanagement._2011_06.VariantManagement.BOMVariantConfigOptionResponse;
import com.teamcenter.services.internal.rac.structuremanagement._2011_06.VariantManagement.BOMVariantConfigOutput;
public class MFCUtility {
	public static void setSWTCenter(Shell shell)
	{
		Rectangle bounds = Display.getDefault().getPrimaryMonitor().getBounds();
		Rectangle rect = shell.getBounds();
		int x = bounds.x + (bounds.width - rect.width) / 2;
		int y = bounds.y + (bounds.height - rect.height) / 2;
		shell.setLocation(x, y);
	}
	public static void stopEdit(JTable table) {
		if(table != null && table.getRowCount() == 0){
			return;
		}
		if (table != null && table.getCellEditor() != null) {
			table.getCellEditor().stopCellEditing();
		}
	}
	public static boolean checkPrivilege(TCComponent component, String privilge) {
		TCAccessControlService service = ((TCSession)AIFUtility.getDefaultSession()).getTCAccessControlService();
		boolean bool = true;
		try {
			bool = service.checkPrivilege(component, privilge);
		} catch (TCException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
			bool = false;
		}
		return bool;
	}
	public static void copy2ClipboardText(String writeMe)
	{
		Clipboard clip = Toolkit.getDefaultToolkit().getSystemClipboard();
		Transferable tText = new StringSelection(writeMe);
		clip.setContents(tText, null);
	}
	public static void warningMassges(String info){
		final String warning = info;
		Display.getDefault().asyncExec(new Runnable() {
			public void run() {
				Shell shell = AIFUtility.getActiveDesktop().getShell();
				MessageDialog.openWarning(shell, "提示",  warning);
			}
		});
	}
	public static void infoMassges(String info){
		final String warning = info;
		Display.getDefault().asyncExec(new Runnable() {
			public void run() {
				Shell shell = AIFUtility.getActiveDesktop().getShell();
				MessageDialog.openInformation(shell, "提示",  warning);
			}
		});
	}
	public static void errorMassges(String info){
		final String warning = info;
		Display.getDefault().asyncExec(new Runnable() {
			public void run() {
				Shell shell = AIFUtility.getActiveDesktop().getShell();
				MessageDialog.openError(shell, "提示",  warning);
			}
		});
	}
	public static  int transA2num(String str){
        int i = 0 ;
        String temp = str.toUpperCase();
        if(str.length()==1){
            i = temp.charAt(0)-'A';
            if(i < 0){
            	i = Integer.parseInt(str) - 1;//转行
            }
        }else{
            i = (temp.charAt(0)-'A' +1 )*26 + temp.charAt(1)-'A';
        }
        return i;
    }
	public static TCComponentFolder createFolder(TCSession session, String name , String desc, String type){
		TCComponentFolder folder = null;
		try{
			DataManagement.CreateIn[] inputs = new DataManagement.CreateIn[1];
			inputs[0] = new DataManagement.CreateIn();
			inputs[0].data.boName = type;
			inputs[0].data.stringProps = new HashMap<String, String>();
			inputs[0].data.stringProps.put("object_name", name);
			inputs[0].data.stringProps.put("object_desc", desc);
			DataManagementService datamanagementservice = DataManagementService.getService(session);
			DataManagement.CreateResponse response = datamanagementservice.createObjects(inputs);
			DataManagement.CreateOut[] out = response.output;
			if(out != null && out.length > 0){
				folder = (TCComponentFolder)out[0].objects[0];
			}
		}catch(Exception e){
			e.printStackTrace();
		}
		return folder;
	}
	public static boolean openByPass(TCSession session) {
		try {
			TCUserService userservice = session.getUserService();
			Object[] objs = { 1 };
			userservice.call("openByPass", objs);
			return true;
		} catch (Exception ex) {
//			MessageBox
//					.post("调用“openByPass”出错!", "提示：", MessageBox.WARNING);
			ex.printStackTrace();
		}
		return false;
	}
	public static boolean closeByPass(TCSession session) {
		try {
			TCUserService userservice = session.getUserService();
			Object[] objs = { 0 };
			userservice.call("closeByPass", objs);
			return true;
		} catch (Exception ex) {
//			MessageBox
//					.post("调用“openByPass”出错!", "提示：", MessageBox.WARNING);
			ex.printStackTrace();
		}
		return false;
	}
	public static TCComponentItemRevision[] getTopRevs(TCComponentItemRevision cRev) {
		TCComponentItemRevision[] parents = null;
		try {
			TCComponent[] comps = (TCComponent[]) cRev.getSession().getUserService().call("getTopItem", new Object[] {cRev});
			System.out.println("comps.len := " + comps.length);
			List<TCComponentItemRevision> lstRevs = new ArrayList<TCComponentItemRevision>();
			for(int i = 0; i < comps.length ; i ++) {
				if(comps[i] instanceof TCComponentItemRevision) {
					lstRevs.add((TCComponentItemRevision)comps[i]);
				}
			}
			System.out.println("lstRevs.size := " + lstRevs.size());
			parents = lstRevs.toArray(new TCComponentItemRevision[lstRevs.size()]);
		}catch (Exception ex) {
			ex.printStackTrace();
		}
		return parents;
	}
	public static TCComponentTask getCurrentTask4Component(TCComponent comp) {
		TCComponentTask task = null;
		try {
			TCComponent[] tasks = comp.getReferenceListProperty("process_stage_list");
			if(tasks != null && tasks.length >0) {
				int i = 0; 
				int len = tasks.length;
				for(i = 0; i < len; i ++) {
					TCComponentTask lst = (TCComponentTask)tasks[i];
					TCComponentTask root = lst.getRoot();
					if(lst != root) {
						task = lst;
						break;
					}
				}
			}
		}catch(Exception e) {
			e.printStackTrace();
		}
		return task;
	}
	public static String[][] getCommonLOVValues(String lovName){
		String[][] retValues = null;
		try {
			TCComponentListOfValues lov = TCComponentListOfValuesType.findLOVByName(lovName);
			String[] dispaly = lov.getListOfValues().getLOVDisplayValues();
			String[] values = lov.getListOfValues().getStringListOfValues();
			retValues = new String[][]{values, dispaly};
		}catch(Exception e) {
			e.printStackTrace();
		}
		return retValues;
	}
	public static TCComponent[] queryComponents(String queryName, String[] clause, String[] values) {
		TCComponent[] comps = null;
		try {
			TCSession session = (TCSession)AIFUtility.getDefaultSession();
			TCComponentQueryType queryType = (TCComponentQueryType)session.getTypeComponent("ImanQuery");
			TCComponentQuery query = (TCComponentQuery)queryType.find(queryName);
			comps = query.execute(clause, values);
		}catch(Exception e) {
			e.printStackTrace();
		}
		return comps;
	}
    /**
     *1. 从剪切板获得文字。
     */
    public static String getSysClipboardText() {
        String ret = "";
        Clipboard sysClip = Toolkit.getDefaultToolkit().getSystemClipboard();
        // 获取剪切板中的内容
        Transferable clipTf = sysClip.getContents(null);

        if (clipTf != null) {
            // 检查内容是否是文本类型
            if (clipTf.isDataFlavorSupported(DataFlavor.stringFlavor)) {
                try {
                    ret = (String) clipTf
                            .getTransferData(DataFlavor.stringFlavor);
                } catch (Exception e) {
                    e.printStackTrace();
                }
            }
        }

        return ret;
    }

    /**
     * 2.将字符串复制到剪切板。
     */
    public static void setSysClipboardText(String writeMe) {
        Clipboard clip = Toolkit.getDefaultToolkit().getSystemClipboard();
        Transferable tText = new StringSelection(writeMe);
        clip.setContents(tText, null);
    }
    public static void stopTableEditing(JTable table) {
		if (table.getCellEditor() != null)
			table.getCellEditor().stopCellEditing();
		table.editingStopped(null);
	}
    public static String readTextFile(File file) {
    	StringBuffer sbText = new StringBuffer();
    	if(file.isFile() && file.exists()) {
    		try {
				InputStreamReader isr = new InputStreamReader(new FileInputStream(file));
				BufferedReader br = new BufferedReader(isr);
			      String lineTxt = null;
			      while ((lineTxt = br.readLine()) != null) {
			        System.out.println(lineTxt);
			        sbText.append(lineTxt);
			      }
			      br.close();
			} catch (Exception e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			}
    	}
    	return sbText.toString();
    }
    public static TCComponentBOMLine[] getWholeBOMLines(TCComponentBOMLine pline) {
    	TCComponentBOMLine[] bomlines = null;
    	try {
			TCComponent[] comps = (TCComponent[]) pline.getSession().getUserService().call("cust_get_bomlines", new Object[] {pline});
			System.out.println("comps.len := " + comps.length);
			List<TCComponentBOMLine> lstRevs = new ArrayList<TCComponentBOMLine>();
			for(int i = 0; i < comps.length ; i ++) {
				if(comps[i] instanceof TCComponentBOMLine) {
					lstRevs.add((TCComponentBOMLine)comps[i]);
				}
			}
			System.out.println("lstRevs.size := " + lstRevs.size());
			bomlines = lstRevs.toArray(new TCComponentBOMLine[lstRevs.size()]);
		}catch (Exception ex) {
			ex.printStackTrace();
		}
    	return bomlines;
    }
    public static String fileNameReplace(String str, String replaceStr) throws PatternSyntaxException {
		// String
		// regEx="[`~!@#$%^&*()+=|{}':;',\\[\\].<>/?~！@#￥%……&*（）——+|{}【】‘；：”“’。，、？]";
		String regEx = "[*|':\\[\\]<>/?*：[\"]”“’？]";
		Pattern p = Pattern.compile(regEx);
		Matcher m = p.matcher(str);
		String newString = m.replaceAll(replaceStr).trim();
		Pattern p2 = Pattern.compile("\\s*|\t|\r|\n");
		Matcher m2 = p2.matcher(newString);
		return m2.replaceAll("").trim();
	}
    public static String transLine2Body(String line) {
    	if((line.contains("01") && line.contains("EC"))
    			|| (line.contains("02") && line.contains("FF"))
    			|| (line.contains("03") && line.contains("RF"))
    			|| (line.contains("FM") && line.contains("FM"))) {
    		return "下屋";
    	}else  if((line.contains("05")  && line.contains("BS"))
    			|| (line.contains("06") && line.contains("ROOF") )
    			|| (line.contains("07") && line.contains("BM"))) {
    		return "上屋";
    	}else  if(line.contains("08") && line.contains("COVER") ) {
    		return "COVER";
    	}else  if(line.contains("09") && line.contains("METAL")) {
    		return "METAL";
    	}
    	return "";
    }
    public static Map<String, List<String>> getVariantCondition(TCComponentBOMLine line){
    	Map<String, List<String>> map= new HashMap<String, List<String>>();
    	try {
    		TCComponent conditionTag = line.getReferenceProperty("bl_condition_tag");
    		if(conditionTag != null) {
    			VariantClause clause;
    			VariantOption option ;
    			String name ;
    			String value ;
    			VariantCondition varCond2 = VariantCondition.create(conditionTag, line.window() );
    			for (int j = 0; j < varCond2.size(); j++) {
    				clause = varCond2.details(j);
					option = clause.getOption();
					if(option==null)
					{
						continue;
					}
					name = option.askName();
					value = clause.getValue();
					System.out.println("选项："+name+" 值："+value);
					if(map.containsKey(name)) {
						List<String> lstValues = map.get(name);
						if(!lstValues.contains(name)) {
							lstValues.add(value);
							map.put(name, lstValues);
						}
					}else {
						List<String> lstValues = new ArrayList<String>();
						lstValues.add(value);
						map.put(name, lstValues);
					}
    			}
    		}
    	}catch(Exception e) {
    		e.printStackTrace();
    	}
    	return map;
    }
    public static String getVariantConditions(TCComponentBOMLine line){
    	Map<String, List<String>> mapVeh = MFCUtility.getVariantCondition(line);
		List<String> lstVehs = mapVeh.get("veh");
		StringBuffer sbVec = new StringBuffer();
		if(lstVehs != null) {
			System.out.println("lstVehs.size ;= " + lstVehs.size());
			
			for(int m = 0; m < lstVehs.size() ; m ++	) {
				if(sbVec.toString().length() > 0) {
					sbVec.append(",");
				}
				sbVec.append(lstVehs.get(m));
			}
		}
		return sbVec.toString();
    }
    public static String getEquipmentByJobName(String jobContent){
    	StringBuffer sbEquipment = new StringBuffer();
    	if(!StringUtil.isEmpty(jobContent)) {
    		if(jobContent.contains("PSW")) {
    			if(sbEquipment.toString().length() > 0) {
    				sbEquipment.append("/");
    			}
    			sbEquipment.append("PSW点焊机");
    		}
    		if(jobContent.contains("RSW")) {
    			if(sbEquipment.toString().length() > 0) {
    				sbEquipment.append("/");
    			}
    			sbEquipment.append("RSW点焊机");
    		}
    		if(jobContent.contains("MSW")) {
    			if(sbEquipment.toString().length() > 0) {
    				sbEquipment.append("/");
    			}
    			sbEquipment.append("MSW点焊机");
    		}
    		if(jobContent.contains("弧焊")) {
    			if(sbEquipment.toString().length() > 0) {
    				sbEquipment.append("/");
    			}
    			sbEquipment.append("弧焊机");
    		}
    		if(jobContent.contains("螺栓焊") || jobContent.contains("螺母焊")) {
    			if(sbEquipment.toString().length() > 0) {
    				sbEquipment.append("/");
    			}
    			sbEquipment.append("固定点焊机");
    		}
            if(jobContent.contains("螺柱焊")) {
    			if(sbEquipment.toString().length() > 0) {
    				sbEquipment.append("/");
    			}
    			sbEquipment.append("螺柱焊机");
    		}
            if(jobContent.contains("激光焊")) {
    			if(sbEquipment.toString().length() > 0) {
    				sbEquipment.append("/");
    			}
    			sbEquipment.append("激光熔焊头或激光纤焊头");
    		}
            if(jobContent.contains("涂胶")) {
    			if(sbEquipment.toString().length() > 0) {
    				sbEquipment.append("/");
    			}
    			sbEquipment.append("手动涂胶枪或自动涂胶枪");
    		}
            if(jobContent.contains("装配")) {
    			if(sbEquipment.toString().length() > 0) {
    				sbEquipment.append("/");
    			}
    			sbEquipment.append("装配工具");
    		}
            if(jobContent.contains("拉铆")) {
    			if(sbEquipment.toString().length() > 0) {
    				sbEquipment.append("/");
    			}
    			sbEquipment.append("铆钉枪和拉铆枪");
    		}
            if(jobContent.contains("打刻")) {
    			if(sbEquipment.toString().length() > 0) {
    				sbEquipment.append("/");
    			}
    			sbEquipment.append("打刻机");
    		}
            if(jobContent.contains("HEM")) {
    			if(sbEquipment.toString().length() > 0) {
    				sbEquipment.append("/");
    			}
    			sbEquipment.append("压边机和滚边机");
    		}
            if(jobContent.contains("打孔")) {
    			if(sbEquipment.toString().length() > 0) {
    				sbEquipment.append("/");
    			}
    			sbEquipment.append("激光打孔机及机械打孔机");
    		}
    	}
    	return sbEquipment.toString();
    }
    public static String getMgrItemsByJobName(String jobContent){
    	StringBuffer sbItems = new StringBuffer();
    	if(!StringUtil.isEmpty(jobContent)) {
    		if(jobContent.contains("点焊") || jobContent.contains("PSW")
    				|| jobContent.contains("RSW") || jobContent.contains("临时参数")) {
    			if(sbItems.toString().length()> 0) {
    				sbItems.append("、");
    			}
    			sbItems.append("A");
    		}
    		if(jobContent.contains("螺母焊")) {
    			if(sbItems.toString().length()> 0) {
    				sbItems.append("、");
    			}
    			sbItems.append("B");
    		}
    		if(jobContent.contains("螺栓焊") || jobContent.contains("螺柱焊")) {
    			if(sbItems.toString().length()> 0) {
    				sbItems.append("、");
    			}
    			sbItems.append("C");
    		}
    		if(jobContent.contains("弧焊") || jobContent.contains("MIG焊") || jobContent.contains("MAG焊")) {
    			if(sbItems.toString().length()> 0) {
    				sbItems.append("、");
    			}
    			sbItems.append("D");
    		}
    		if(jobContent.contains("安装")) {
    			if(sbItems.toString().length()> 0) {
    				sbItems.append("、");
    			}
    			sbItems.append("E");
    		}
    		if(jobContent.contains("涂胶")) {
    			if(sbItems.toString().length()> 0) {
    				sbItems.append("、");
    			}
    			sbItems.append("F");
    		}
    		if(jobContent.contains("门盖总成") || jobContent.contains("外观检查") || jobContent.contains("建付调整")) {
    			if(sbItems.toString().length()> 0) {
    				sbItems.append("、");
    			}
    			sbItems.append("G");
    		}
    		if(jobContent.contains("前门装配") || jobContent.contains("后门装配") || jobContent.contains("机盖装配") || 
    				jobContent.contains("行李箱盖装配") || jobContent.contains("前翼子板装配") || jobContent.contains("3D测量")
    				|| jobContent.contains("建付测量")) {
    			if(sbItems.toString().length()> 0) {
    				sbItems.append("、");
    			}
    			sbItems.append("H");
    		}
    		if(jobContent.contains("点焊") || jobContent.contains("VIN打刻") || jobContent.contains("涂胶")) {
    			if(sbItems.toString().length()> 0) {
    				sbItems.append("、");
    			}
    			sbItems.append("I");
    		}
    		if(jobContent.contains("前后门") || jobContent.contains("机盖") || jobContent.contains("HEM")) {
    			if(sbItems.toString().length()> 0) {
    				sbItems.append("、");
    			}
    			sbItems.append("J");
    		}
    		if(jobContent.contains("构成表") ) {
    			if(sbItems.toString().length()> 0) {
    				sbItems.append("、");
    			}
    			sbItems.append("K");
    		}
    		if(jobContent.contains("拉铆") ) {
    			if(sbItems.toString().length()> 0) {
    				sbItems.append("、");
    			}
    			sbItems.append("L");
    		}
    		if(jobContent.contains("激光焊接") ) {
    			if(sbItems.toString().length()> 0) {
    				sbItems.append("、");
    			}
    			sbItems.append("M");
    		}
    		if(jobContent.contains("激光打孔") ) {
    			if(sbItems.toString().length()> 0) {
    				sbItems.append("、");
    			}
    			sbItems.append("N");
    		}
    		if(jobContent.contains("激光钎焊") ) {
    			if(sbItems.toString().length()> 0) {
    				sbItems.append("、");
    			}
    			sbItems.append("O");
    		}
    		if(jobContent.contains("加强贴贴附") ) {
    			if(sbItems.toString().length()> 0) {
    				sbItems.append("、");
    			}
    			sbItems.append("P");
    		}
    	}
    	return sbItems.toString().length() > 0 ? "B表" + sbItems.toString() + "项" : "";
    }
    public static void loadProperties(TCSession session, TCComponent[] comps, String[] properties) {
		if(comps!=null&&comps.length>0)
		{
			DataManagementService.getService(session).getProperties(comps, properties);
		}
	}
    public static String[] getVariantValues(TCComponentBOMLine line) {
    	System.out.println("getVariantValues");
    	VariantManagementService variantManagementService = VariantManagementService.getService(line.getSession());
    	String[] allvalues = null;
    	try {
    		BOMVariantConfigOptionResponse bOMVariantConfigOptionResponse = 
        			variantManagementService.getBOMVariantConfigOptions(line.window(), line);
            BOMVariantConfigOutput bOMVariantConfigOutput = bOMVariantConfigOptionResponse.output;
            for (int i = 0; i < bOMVariantConfigOutput.configuredOptions.length; i++) 
            {
                  VariantManagement.BOMVariantConfigurationOption bOMVariantConfigurationOption = bOMVariantConfigOutput.configuredOptions[i];
                  System.out.println("bOMVariantConfigurationOption.variantType:"+bOMVariantConfigurationOption.variantType);
                  if ("BOM_LEGACY".equals(bOMVariantConfigurationOption.variantType)) {
                	  String CONFIGURATION =  bOMVariantConfigurationOption.valueSet;
                	  allvalues = bOMVariantConfigurationOption.classicOption.optionValues;
                	  System.out.println("value set:"+CONFIGURATION);
                	  for(int j = 0; j < allvalues.length; j ++) {
                		  System.out.println("option value := " + allvalues[j]);
                	  }
                	  break;
                  }
            }
    	}catch(Exception e) {
    		e.printStackTrace();
    	}
    	
    	return allvalues;
    }
}
