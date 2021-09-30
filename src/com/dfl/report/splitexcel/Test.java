package com.dfl.report.splitexcel;

public class Test {

	public static void main(String[] args) {
		// TODO Auto-generated method stub
//		String s = "工程作业表-14临时参数";
//		System.out.println(s.substring(s.lastIndexOf("-")+1, s.length()));
		
		
	     String str="06assu\\sssme345c\\onssstribute\\";  
			String name = str.replace("\\", "");
			System.out.println("name:"+name);
	     
	     if(!str.contains("\\")||str.endsWith("\\")	)
	     {
	    	 System.out.println("111");
	     }
//	     System.out.println(str.replaceFirst("\\d+",""));  
		
	}

}
