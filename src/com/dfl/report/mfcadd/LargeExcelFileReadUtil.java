package com.dfl.report.mfcadd;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.Iterator;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import java.util.Map.Entry;

import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.xssf.eventusermodel.XSSFReader;
import org.apache.poi.xssf.model.SharedStringsTable;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;
import org.xml.sax.Attributes;
import org.xml.sax.ContentHandler;
import org.xml.sax.InputSource;
import org.xml.sax.SAXException;
import org.xml.sax.XMLReader;
import org.xml.sax.helpers.DefaultHandler;
import org.xml.sax.helpers.XMLReaderFactory;
/**
 * �������Ƚϴ�(8��������)��excel�ļ���������excel�ļ�����Ϊ ��������-ֵ����ʽ����map�У��˷�ʽ�ٶȿ죬�ڴ����С ��ֻ�ܶ�ȡexcle�ļ�
 * �ṩ������sheet���� processOneSheet(String  filename) �Լ�������sheet���� processAllSheets(String  filename)
 * ֻ�贫���ļ�·��+�ļ�������  ���ô�����������ֻ�� ����LargeExcelFileReadUtil.getRowContents()����ֵ���ɻ�ý����������
 *
 */
public class LargeExcelFileReadUtil  {

    private  LinkedHashMap<String, String>rowContents=new LinkedHashMap<String, String>(); 
    private  SheetHandler sheetHandler;

public LinkedHashMap<String, String> getRowContents() {
        return rowContents;
    }
    public void setRowContents(LinkedHashMap<String, String> rowContents) {
        this.rowContents = rowContents;
    }

    public SheetHandler getSheetHandler() {
        return sheetHandler;
    }
    public void setSheetHandler(SheetHandler sheetHandler) {
        this.sheetHandler = sheetHandler;
    }
    //����һ��sheet
    public void processOneSheet(String filename) throws Exception {
        InputStream sheet2=null;
        OPCPackage pkg =null;
        try {
                pkg = OPCPackage.open(filename);
                XSSFReader r = new XSSFReader(pkg);
                SharedStringsTable sst = r.getSharedStringsTable();
                XMLReader parser = fetchSheetParser(sst);
                sheet2 = r.getSheet("rId1");
                InputSource sheetSource = new InputSource(sheet2);
                parser.parse(sheetSource);
                setRowContents(sheetHandler.getRowContents());
                }catch (Exception e) {
                    e.printStackTrace();
                    throw e;
                    }finally{
                        if(pkg!=null){
                            pkg.close();
                                     }
                        if(sheet2!=null){
                            sheet2.close();
                                        }
                }
    }
//������sheet
    public void processAllSheets(String filename) throws Exception {
        OPCPackage pkg =null;
        InputStream sheet=null;
        try{
                pkg=OPCPackage.open(filename);
                XSSFReader r = new XSSFReader( pkg );
                SharedStringsTable sst = r.getSharedStringsTable();
                XMLReader parser = fetchSheetParser(sst);
                Iterator<InputStream> sheets = r.getSheetsData();
                while(sheets.hasNext()) {
                    System.out.println("Processing new sheet:\n");
                    sheet = sheets.next();
                    InputSource sheetSource = new InputSource(sheet);
                    parser.parse(sheetSource);
                                        }
            }catch (Exception e) {
                    e.printStackTrace();
                    throw e;
                   }finally{
                       if(pkg!=null){
                           pkg.close();
                                 }
                       if(sheet!=null){
                           sheet.close();
                                    }
                            }
    }

    public XMLReader fetchSheetParser(SharedStringsTable sst) throws SAXException {
        XMLReader parser =
            XMLReaderFactory.createXMLReader(
                    "com.sun.org.apache.xerces.internal.parsers.SAXParser"
            );
        setSheetHandler(new SheetHandler(sst));
        ContentHandler handler = (ContentHandler) sheetHandler;
        parser.setContentHandler(handler);
        return parser;
    }
    static public int transA2num(String str){
        int i = 0 ;
        String temp = str.toUpperCase();
        if(str.length()==1){
            i = temp.charAt(0)-'A';
        }else{
            i = (temp.charAt(0)-'A' +1 )*26 + temp.charAt(1)-'A';
        }
        return i;
    }
    public String[][] getExcelDatas(String xlsxPath){
 	   Long time = System.currentTimeMillis();
 	   try {
 		processOneSheet(xlsxPath);
 		LinkedHashMap<String, String>  mapDatas = getRowContents();
 		LinkedHashMap<String, String>  mapDatas2 = getRowContents();
 		Iterator<Entry<String, String>> it = mapDatas.entrySet().iterator();
 		 int columns = 0;
 		 int rows = 0;
 		 while(it.hasNext()) {
 			 Map.Entry<String, String> entry=(Map.Entry<String, String>)it.next();
 	         String pos = entry.getKey(); 
 	         int col = 0;
 	         String row = "";
 	         if(pos.charAt(1) >= 'A' && pos.charAt(1) <= 'Z') {
 	        	 col = transA2num(pos.substring(0, 2));
 	        	 row = pos.substring(2);
 	         }else {
 	        	 col = transA2num(pos.substring(0, 1));
 	        	 row = pos.substring(1);
 	         }
 	         if(Integer.parseInt(row) > rows) {
 	        	 rows = Integer.parseInt(row);
 	         }
 	         if(col > columns) {
 	        	 columns = col;
 	         }
 		 }
 		 columns ++;
 		 System.out.println("columns := " + columns);
 		 String[][] rowdatas = new String[rows][columns];
 		 it = mapDatas2.entrySet().iterator();
 		 while(it.hasNext()) {
 			 Map.Entry<String, String> entry=(Map.Entry<String, String>)it.next();
 	         String pos = entry.getKey(); 
 	         String value = entry.getValue();
 	         
 	         int col = 0;
 	         String row = "";
 	         if(pos.charAt(1) >= 'A' && pos.charAt(1) <= 'Z') {
 	        	 col = transA2num(pos.substring(0, 2));
 	        	 row = pos.substring(2);
 	         }else {
 	        	 col = transA2num(pos.substring(0, 1));
 	        	 row = pos.substring(1);
 	         }
 	         rowdatas[Integer.parseInt(row) - 1][col] = value;
 		 }
 		 Long endtime = System.currentTimeMillis();
 		 System.out.println("��ȡExcel��ʱ"+(endtime-time)/1000+"��");
 		 //System.out.println(rowdatas[rows -1][columns-1]);
 		 return rowdatas;
 	} catch (Exception e) {
 		// TODO Auto-generated catch block
 		e.printStackTrace();
 	}
 	   return null;
    }
    /** 
     * See org.xml.sax.helpers.DefaultHandler javadocs 
     */
    //����
    public static void test ()throws Exception {
       Long time=System.currentTimeMillis();
        LargeExcelFileReadUtil example = new LargeExcelFileReadUtil();

        example.processOneSheet("E:/2.xlsx");
        Long endtime = System.currentTimeMillis();
        LinkedHashMap<String, String>  map= example.getRowContents();
        Iterator<Entry<String, String>> it= map.entrySet().iterator();
        int count=0;
        String prePos="";
        while (it.hasNext()){
            Map.Entry<String, String> entry=(Map.Entry<String, String>)it.next();
            String pos = entry.getKey();
            String value = entry.getValue();
            System.out.println("pos := " + pos + " --> " + value);
            if(!pos.substring(1).equals(prePos)){
                prePos=pos.substring(1);
                count++;
            }
            System.out.println(pos+";"+entry.getValue());
        }
        System.out.println("��������"+count+"��;��ʱ"+(endtime-time)/1000+"��");
    }
    public static void main(String[] args) {
    	try {
    		 LargeExcelFileReadUtil example = new LargeExcelFileReadUtil();
    		 String[][] infos = example.getExcelDatas("E:\\1.xlsx");
    		 for(int i = 0; i < infos.length; i ++) {
    			 System.out.println("info := " + infos[i][1]);
    		 }
    		// List<String[]> lst = example.getExcelDatas("E:/1.xlsx");
		} catch (Exception e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
    }
}