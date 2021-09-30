package com.dfl.report.splitexcel;

import java.io.File;
import java.io.IOException;
import java.util.Date;

import org.apache.log4j.lf5.util.ResourceUtils;

import com.dfl.report.util.Util;

public class SplitExcelTest {

	public SplitExcelTest() {
		// TODO Auto-generated constructor stub
	}

	public static void main(String[] args) {
		// TODO Auto-generated method stub
		File scriptFile = Util.getRCPPluginInsideFile("CopySheet.vbs");
		
		if(scriptFile!=null)
		{
			String xlsFilePath="C:\\Users\\Administrator\\Desktop\\gw.xlsx";
			String oupFilePath = "C:\\Users\\Administrator\\Desktop\\测试.xlsx";
			
			File testfile = new File("C:\\Users\\Administrator\\Desktop\\gw.xlsx");
			String name = testfile.getName();
			
			String tempPath = getTempPath();
			//String oupFilePath = tempPath+"output";
//			File dirfile = new File(oupFilePath);
//			if(!dirfile.exists())
//			{
//				dirfile.mkdir();
//			}
			final String command = "wscript  \"" + scriptFile.getAbsolutePath() + "\" \"" + xlsFilePath + "\" \"" + oupFilePath + "\" \"03构成表1" + "\" \"03A构成表1\"";
			
			System.out.println(command);
			try {
				Process	process = Runtime.getRuntime().exec(command);
				try {
					process.waitFor();
				} catch (InterruptedException e) {
					// TODO Auto-generated catch block
					e.printStackTrace();
				}
				System.out.println("finish");
				
				String oupFilePath2 = tempPath+"outputsuccess";

				File file = new File(oupFilePath2);
				if(file.exists())
				{
					File[] files = file.listFiles();
					for (int i = 0; i < files.length; i++) {
						System.out.println("files:"+files[i].getPath());
					}
				}else
				{
					System.out.println("vbs查房");
				}
				
			} catch (IOException e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			}
		}
		
	}
	
	
	public static String getTempPath()
	{
		String path = "";
		String tmpPath = System.getProperty("java.io.tmpdir");
//		System.out.println("tmpPath:"+tmpPath);
		if(tmpPath.endsWith("\\"))
		{
			path = tmpPath+new Date().getTime();
		}else
		{
			path = tmpPath+"\\"+new Date().getTime();
		}
		path= path+"\\";
		System.out.println("tempPath="+path);
		
		File dirfile = new File(path);
		if(!dirfile.exists())
		{
			dirfile.mkdir();
		}
		return path;
	}
	

}
