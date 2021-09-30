package com.dfl.report.workschedule;

import java.util.ArrayList;
import java.util.List;

public class MoveUtil {

	
	public static ArrayList moveUp(ArrayList list,int startIndex,int endIndex) {
		// TODO Auto-generated method stub
		if(startIndex==0)
		{
			return list;
		}
		List tempList = list.subList(startIndex, endIndex+1);
		ArrayList colneList = new ArrayList(tempList);
		list.subList(startIndex, endIndex+1).clear();
		list.removeAll(colneList);
		list.addAll(startIndex-1, colneList);
		return list;
	}
	
	public static ArrayList moveDown(ArrayList list,int startIndex,int endIndex) {
		// TODO Auto-generated method stub
		if(endIndex==(list.size()-1))
		{
			return list;
		}
		List tempList = list.subList(startIndex, endIndex+1);
		ArrayList colneList = new ArrayList(tempList);
		list.subList(startIndex, endIndex+1).clear();
		list.removeAll(colneList);
		
		
		list.addAll(startIndex+1, colneList);

		
		return list;
	}

}
