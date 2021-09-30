package com.dfl.report.workschedule;

import java.util.List;

import org.eclipse.jface.viewers.CellLabelProvider;
import org.eclipse.jface.viewers.ICellModifier;
import org.eclipse.jface.viewers.ILabelProviderListener;
import org.eclipse.jface.viewers.IStructuredContentProvider;
import org.eclipse.jface.viewers.ITableColorProvider;
import org.eclipse.jface.viewers.ITableLabelProvider;
import org.eclipse.jface.viewers.TableViewer;
import org.eclipse.jface.viewers.Viewer;
import org.eclipse.jface.viewers.ViewerCell;
import org.eclipse.swt.graphics.Color;
import org.eclipse.swt.graphics.Image;
import org.eclipse.swt.widgets.TableItem;


public class TableProvider1 extends CellLabelProvider implements ITableLabelProvider,ITableColorProvider, IStructuredContentProvider ,ICellModifier {
	public TableViewer tableViewer;
	private String[] proNames = null;
	public TableProvider1(TableViewer tableViewer,String[] proNames) {
		this.tableViewer = tableViewer;
		this.proNames = proNames;
	}
	
	@Override
	public void dispose() {

	}
	@Override
	public void inputChanged(Viewer viewer, Object oldInput, Object newInput) {
		// TODO Auto-generated method stub
//		tableViewer = (TableViewer) viewer;
		
	}

	@Override
	public boolean canModify(Object element, String property) {
		TableInfo info = (TableInfo) element;
		if(info.isCanEdit())
		{
			return true;
		}
		return false;
	}

	@Override
	public Object[] getElements(Object input) {
		if(input ==null){
			return new Object[0];
		}
		if(input instanceof Object[]){
			return (Object[]) input;
		}else if(input instanceof List){
			List<Object> list = (List<Object>) input;
			return list.toArray(new Object[list.size()]);
		}
		return new Object[0];
	}

	@Override
	public void addListener(ILabelProviderListener listener) {
		// TODO Auto-generated method stub
		
	}

	@Override
	public boolean isLabelProperty(Object element, String property) {
		// TODO Auto-generated method stub
		return false;
	}

	@Override
	public void removeListener(ILabelProviderListener listener) {
		// TODO Auto-generated method stub
		
	}

	@Override
	public Image getColumnImage(Object element, int columnIndex) {
		if(element==null){
			return null;
		}
		
		Image image = null;
		return image;
	}

	@Override
	public String getColumnText(Object element, int columnIndex) {
		if(element==null){
			return "";
		}
		TableInfo row = (TableInfo) element;
		return row.getValue(proNames[columnIndex]);
	}


	@Override
	public void modify(Object element, String property, Object value) {
		// TODO Auto-generated method stub
		if(property.equals("page"))
		{
			
			String str = value.toString();
			try
			{
				Integer.parseInt(str);
			}catch (NumberFormatException e) {
				return;
			}
			
			TableItem item  = (TableItem) element;
			TableInfo info = (TableInfo) item.getData();
			info.setPage(str);
			tableViewer.update(info, null);
		}
		
		return;
	}


	@Override
	public Color getForeground(Object paramObject, int paramInt) {
		// TODO Auto-generated method stub
		return null;
	}
	

	@Override
	public Color getBackground(Object paramObject, int paramInt) {
		// TODO Auto-generated method stub
		return null;
	}


	@Override
	public void update(ViewerCell paramViewerCell) {
		// TODO Auto-generated method stub
		
	}

	@Override
	public Object getValue(Object paramObject, String proName) {
		if(paramObject==null){
			return "";
		}
		TableInfo row = (TableInfo) paramObject;
		return row.getValue(proName);
	}



}
