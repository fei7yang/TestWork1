package com.dfl.report.home;

import java.util.ArrayList;

import org.apache.log4j.Logger;
import org.eclipse.jface.viewers.CellLabelProvider;
import org.eclipse.jface.viewers.ICellModifier;
import org.eclipse.jface.viewers.ITableColorProvider;
import org.eclipse.jface.viewers.ITableLabelProvider;
import org.eclipse.jface.viewers.ITreeContentProvider;
import org.eclipse.jface.viewers.TreeViewer;
import org.eclipse.jface.viewers.Viewer;
import org.eclipse.jface.viewers.ViewerCell;
import org.eclipse.swt.graphics.Color;
import org.eclipse.swt.graphics.Image;

import com.teamcenter.rac.common.TCTypeRenderer;
import com.teamcenter.rac.kernel.TCComponent;
import com.teamcenter.rac.kernel.TCComponentFolder;
import com.teamcenter.rac.kernel.TCException;

public class TreeProvider extends CellLabelProvider
implements ITreeContentProvider, ITableLabelProvider, ICellModifier ,ITableColorProvider {
	//IStructuredContentProvider

	private TreeViewer treeViewer;
	static  Logger logger = Logger.getLogger(TreeProvider.class.getName());
	public TreeProvider(TreeViewer treeViewer) {
		// TODO Auto-generated constructor stub
		this.treeViewer = treeViewer;
	}

	@Override
	public void dispose() {
		// TODO Auto-generated method stub

	}

	@Override
	public void inputChanged(Viewer viewer, Object oldInput, Object newInput) {
		// TODO Auto-generated method stub

	}

	@Override
	public Object[] getElements(Object inputElement) {
		// TODO Auto-generated method stub
		// return null;
		return getChildren(inputElement);
	}

	@Override
	public Object[] getChildren(Object parent) {
		// TODO Auto-generated method stub
		// return null;
		TreeNode parentNode = (TreeNode) parent;
		TCComponent parentFolder = parentNode.getFolder();
		ArrayList parentFolderList = getChilds(parentFolder);
		TreeNode[] childs = (TreeNode[]) parentFolderList.toArray(new TreeNode[parentFolderList.size()]);
		return childs;
	}



	private ArrayList getChilds(TCComponent parent) {
		// TODO Auto-generated method stub
		ArrayList list = new ArrayList();
		try {
			TCComponent[] contents = parent.getReferenceListProperty("contents");
			for (int i = 0; i < contents.length; i++) {
				if(contents[i] instanceof TCComponentFolder)
				{
					TreeNode child = new TreeNode(contents[i]);
					list.add(child);
				}
			}
		
		} catch (TCException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		
		return list;
	}

	@Override
	public Object getParent(Object element) {
		// TODO Auto-generated method stub
		return null;
	}

	@Override
	public boolean hasChildren(Object parent) {
		// TODO Auto-generated method stub
		return true;
	}

	@Override
	public void update(ViewerCell cell) {
		// TODO Auto-generated method stub
	}

	@Override
	public Image getColumnImage(Object element, int columnIndex) {
		// TODO Auto-generated method stub
//
//		Image image = SWTResourceManager.getImage(TreeProvider.class,
//				"/com/hanhe/bin/dialog/foldertype_16.png");
		
		Image image = null;
			if(columnIndex == 0)
	        {
				TreeNode row = (TreeNode) element;
	        	 TCComponent comp = row.getFolder();
	             image = TCTypeRenderer.getImage(comp);
	        }
			
			return image;
	}

	@Override
	public String getColumnText(Object element, int columnIndex) {
		// TODO Auto-generated method stub
		TreeNode node = (TreeNode) element;
		//"title","gzlx","wxjb","zl"
		if (columnIndex == 0) {
			return node.getName();
		}
		return null;
	}

	@Override
	public boolean canModify(Object element, String property) {
		// TODO Auto-generated method stub
		return false;
	}

	@Override
	public Object getValue(Object element, String property) {
		// TODO Auto-generated method stub
		TreeNode node = (TreeNode) element;
		String value = "";
		if (property.equals("name")) {
			value = node.getName();
		}
		return value;
	}

	@Override
	public void modify(Object element, String property, Object value) {
		// TODO Auto-generated method stub

	}

	@Override
	public Color getForeground(Object element, int columnIndex) {
		// TODO Auto-generated method stub
		return null;
	}

	@Override
	public Color getBackground(Object element, int columnIndex) {
		// TODO Auto-generated method stub

		return null;
	}


}
