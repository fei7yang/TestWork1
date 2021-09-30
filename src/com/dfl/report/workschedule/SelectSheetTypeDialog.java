package com.dfl.report.workschedule;

import org.eclipse.swt.widgets.Dialog;
import org.eclipse.swt.widgets.Display;
import org.eclipse.swt.widgets.Shell;
import org.eclipse.swt.widgets.Table;
import org.eclipse.swt.widgets.TableColumn;
import org.eclipse.swt.widgets.TableItem;
import org.eclipse.swt.widgets.Composite;

import java.util.ArrayList;
import java.util.HashMap;
import java.util.Iterator;
import java.util.LinkedHashMap;
import java.util.List;

import org.eclipse.jface.viewers.CellEditor;
import org.eclipse.jface.viewers.CellLabelProvider;
import org.eclipse.jface.viewers.IStructuredContentProvider;
import org.eclipse.jface.viewers.LabelProvider;
import org.eclipse.jface.viewers.TableViewer;
import org.eclipse.jface.viewers.TextCellEditor;
import org.eclipse.jface.viewers.Viewer;
import org.eclipse.jface.viewers.ViewerSorter;
import org.eclipse.swt.SWT;
import org.eclipse.swt.widgets.Button;
import org.eclipse.swt.events.SelectionAdapter;
import org.eclipse.swt.events.SelectionEvent;
import org.eclipse.swt.graphics.Image;
import org.eclipse.wb.swt.SWTResourceManager;
import org.omg.CosCollection.Map;

import com.teamcenter.rac.aif.AbstractAIFUIApplication;
import com.teamcenter.rac.aifrcp.AIFUtility;
import com.teamcenter.rac.kernel.TCPreferenceService;
import com.teamcenter.rac.kernel.TCSession;

import swing2swt.layout.BorderLayout;
import org.eclipse.swt.layout.RowLayout;
import swing2swt.layout.FlowLayout;
import org.eclipse.swt.layout.FormLayout;
import org.eclipse.swt.layout.FormData;
import org.eclipse.swt.layout.FormAttachment;
import org.eclipse.swt.layout.GridLayout;
import org.eclipse.swt.widgets.Label;
import org.eclipse.swt.graphics.Point;
import org.eclipse.swt.layout.FillLayout;
import org.eclipse.swt.layout.GridData;
import org.eclipse.swt.widgets.Combo;
import org.eclipse.swt.widgets.Text;

public class SelectSheetTypeDialog extends Dialog {

	private static class ViewerLabelProvider extends LabelProvider {
		public Image getImage(Object element) {
			return super.getImage(element);
		}

		public String getText(Object element) {
			return super.getText(element);
		}
	}

	private static class Sorter extends ViewerSorter {
		public int compare(Viewer viewer, Object e1, Object e2) {
			Object item1 = e1;
			Object item2 = e2;
			return 0;
		}
	}

	private static class ContentProvider implements IStructuredContentProvider {
		public Object[] getElements(Object inputElement) {
			return new Object[0];
		}

		public void dispose() {
		}

		public void inputChanged(Viewer viewer, Object oldInput, Object newInput) {
		}
	}

	protected Object result;
	protected Shell shlsheet;
//	private Table table;
	private TableViewer tableViewer;
	private Combo combo_1;
	public LinkedHashMap<String, String> map = new LinkedHashMap<String, String>();
	public ArrayList list = new ArrayList();
	public String model;
	public String modelname;
	public String sheetname;
	public String sheetpages;
	private Label lblNewLabel;
	private Text text;
	private Text text_1;
	private List shList;

	/**
	 * Create the dialog.
	 * 
	 * @param parent
	 * @param style
	 * @param shList 
	 */
	public SelectSheetTypeDialog(Shell parent, int style, List shList) {
		super(parent, style);
		setText("SWT Dialog");
		this.shList = shList;
	}

	/**
	 * Open the dialog.
	 * 
	 * @return the result
	 */
	public Object open() {
		createContents();
		shlsheet.open();
		shlsheet.layout();
		Display display = getParent().getDisplay();
		centerToScreen(display);
		while (!shlsheet.isDisposed()) {
			if (!display.readAndDispatch()) {
				display.sleep();
			}
		}
		return result;
	}

	protected void centerToScreen(Display display) {
		int nLocationX = display.getClientArea().width / 2 - shlsheet.getSize().x / 2;
		int nLocationY = display.getClientArea().height / 2 - shlsheet.getSize().y / 2;
		shlsheet.setLocation(nLocationX, nLocationY);
	}

	/**
	 * Create contents of the dialog.
	 */
	private void createContents() {
		shlsheet = new Shell(getParent(), getStyle());
		shlsheet.setMinimumSize(new Point(136, 20));
		shlsheet.setImage(SWTResourceManager.getImage(SelectSheetTypeDialog.class,
				"/com/dfl/report/imags/defaultapplication_16.png"));
		shlsheet.setSize(655, 621);
		shlsheet.setText("\u9009\u62E9\u589E\u9875sheet\u7C7B\u578B");
		shlsheet.setLayout(new BorderLayout(0, 0));

		Composite composite = new Composite(shlsheet, SWT.NONE);
		composite.setLayoutData(BorderLayout.NORTH);
		composite.setLayout(null);

		Label lblSheet = new Label(composite, SWT.NONE);
		lblSheet.setAlignment(SWT.RIGHT);
		lblSheet.setBounds(217, 9, 116, 17);
		lblSheet.setText("sheet\u540D\u79F0\uFF1A");

		Label label_1 = new Label(composite, SWT.NONE);
		label_1.setBounds(10, 9, 61, 17);
		label_1.setText("\u6A21\u677F\u7C7B\u578B\uFF1A");

		combo_1 = new Combo(composite, SWT.NONE);
		combo_1.addSelectionListener(new SelectionAdapter() {
			@Override
			public void widgetSelected(SelectionEvent e) {
				SelectloadTableData();
			}
		});
		combo_1.setBounds(72, 6, 139, 25);
		combo_1.add("普通工位模板");
		combo_1.add("VIN码打刻模板");
		combo_1.add("调整线模板");
		combo_1.select(0);
		
		text = new Text(composite, SWT.BORDER);
		text.setBounds(339, 6, 88, 23);
		
		Label lblNewLabel_1 = new Label(composite, SWT.NONE);
		lblNewLabel_1.setAlignment(SWT.RIGHT);
		lblNewLabel_1.setBounds(433, 9, 98, 17);
		lblNewLabel_1.setText("sheet\u9875\u7801\uFF1A");
		
		text_1 = new Text(composite, SWT.BORDER);
		text_1.setBounds(537, 6, 82, 23);

		Composite composite_1 = new Composite(shlsheet, SWT.NONE);
		composite_1.setLayoutData(BorderLayout.SOUTH);
		composite_1.setLayout(null);

		// 确定按钮
		Button btnNewButton = new Button(composite_1, SWT.NONE);
		btnNewButton.setBounds(388, 5, 71, 27);
		btnNewButton.addSelectionListener(new SelectionAdapter() {
			@Override
			public void widgetSelected(SelectionEvent e) {
				ComformEvent();
			}
		});
		btnNewButton.setText("\u786E\u5B9A");

		// 取消按钮
		Button btnNewButton_1 = new Button(composite_1, SWT.NONE);
		btnNewButton_1.setBounds(465, 5, 72, 27);
		btnNewButton_1.addSelectionListener(new SelectionAdapter() {
			@Override
			public void widgetSelected(SelectionEvent e) {
				shlsheet.dispose();
			}
		});
		btnNewButton_1.setText("\u53D6\u6D88");
		
		lblNewLabel = new Label(composite_1, SWT.NONE);
		lblNewLabel.setForeground(SWTResourceManager.getColor(SWT.COLOR_RED));
		lblNewLabel.setBounds(10, 10, 372, 17);

		tableViewer = new TableViewer(shlsheet, SWT.BORDER | SWT.CHECK | SWT.FULL_SELECTION | SWT.MULTI);
		Table table = tableViewer.getTable();
		table.addSelectionListener(new SelectionAdapter() {
			@Override
			public void widgetSelected(SelectionEvent e) {
				SelectSheettype();
			}
		});
		table.setHeaderVisible(true);
		table.setLinesVisible(true);
		table.setLayoutData(BorderLayout.CENTER);

		String tableColumn[] = new String[] { "name", "page" };
		tableViewer.setColumnProperties(tableColumn);// 设置列属性
		TableProvider1 provider_2 = new TableProvider1(tableViewer, tableColumn);
		tableViewer.setContentProvider(provider_2);
		tableViewer.setLabelProvider(provider_2);
		tableViewer.setCellModifier(provider_2);

		TableColumn tblclmnNewColumn = new TableColumn(table, SWT.NONE);
		tblclmnNewColumn.setWidth(402);
		tblclmnNewColumn.setText("Sheet\u540D\u79F0");

		CellEditor acelleditor[] = new CellEditor[2];
		acelleditor[1] = new TextCellEditor(table, 131072);
		tableViewer.setCellEditors(acelleditor);
		

		loadTableData();

	}

	protected void SelectSheettype() {
		// TODO Auto-generated method stub
		Table table = tableViewer.getTable();
		TableItem[] items = table.getItems();// 获取表格所有的行
		List namelist =new ArrayList();
		for (int i = 0; i < items.length; i++) {
			TableItem tableItem = items[i];
			String[] str = new String[2];
			str[0] = tableItem.getText(0);
			if (tableItem.getChecked()) {			
				namelist.add(str[0]);
				text.setText(str[0]);
			}else {
				if(namelist.contains(str[0])) {
					namelist.remove(str[0]);
				}
			}
		}
		if(namelist == null || namelist.size()<1) {
			text.setText("");
		}
	}

	protected void ComformEvent() {
		// TODO Auto-generated method stub
		if (combo_1.getSelectionIndex() == -1) {
			lblNewLabel.setText("模板类型！");
			return;
		}
		if (text.getText().isEmpty()) {
			lblNewLabel.setText("请填写sheet名称！");
			return;
		}
		if(shList.contains(text.getText())) {
			lblNewLabel.setText("报表中已存在该sheet名称，sheet名称不能重复！");
			return;
		}
		if (text_1.getText().isEmpty()) {
			lblNewLabel.setText("请填写sheet页码！");
			return;
		}
		
		model = combo_1.getItem(combo_1.getSelectionIndex());
		sheetname = text.getText();
		sheetpages = text_1.getText();
		Table table = tableViewer.getTable();
		TableItem[] items = table.getItems();// 获取表格所有的行
		for (int i = 0; i < items.length; i++) {
			TableItem tableItem = items[i];
			if (tableItem.getChecked()) {
				String[] str = new String[2];
				str[0] = tableItem.getText(0);
				list.add(str[0]);
			}
		}
		if(list.size()<1) {
			lblNewLabel.setText("请勾选增加的sheet类型！");
			return;
		}
		if(list.size()>1) {
			lblNewLabel.setText("请勾选单个sheet页！");
			list.clear();
			return;
		}
		modelname = (String) list.get(0);
		shlsheet.dispose();
	}

	private void loadTableData() {
		// TODO Auto-generated method stub
		TableInfo[] infos = new TableInfo[25];
		infos[0] = new TableInfo("首页", "", false);
		infos[1] = new TableInfo("有效页", "", false);
		infos[2] = new TableInfo("构成表", "", false);
		infos[3] = new TableInfo("构成图", "", false);
		infos[4] = new TableInfo("式样差", "", false);
		infos[5] = new TableInfo("涂胶", "", false);
		infos[6] = new TableInfo("螺柱焊", "", false);
		infos[7] = new TableInfo("螺母焊", "", false);
		infos[8] = new TableInfo("螺栓焊", "", false);
		infos[9] = new TableInfo("点焊-PSW", "1", true);
		infos[10] = new TableInfo("点焊-RSW", "1", true);
		infos[11] = new TableInfo("点焊-MSW", "1", true);
		infos[12] = new TableInfo("点焊-SSW", "1", true);
		infos[13] = new TableInfo(" PSW", "", false);
		infos[14] = new TableInfo("RSW气动", "", false);
		infos[15] = new TableInfo("RSW伺服", "", false);
		infos[16] = new TableInfo("临时参数", "", false);
		infos[17] = new TableInfo("弧焊作业", "", false);
		infos[18] = new TableInfo("HEM指示", "", false);
		infos[19] = new TableInfo("铰链安装", "", false);
		infos[20] = new TableInfo("装配", "", false);
		infos[21] = new TableInfo("打点统计表", "", false);
		infos[22] = new TableInfo("NEPID安装", "", false);
		infos[23] = new TableInfo("其他", "", false);
		infos[24] = new TableInfo("法兰边检查", "", false);
		tableViewer.setInput(infos);

	}

	private void SelectloadTableData() {

		model = combo_1.getItem(combo_1.getSelectionIndex());
		if (model.equals("普通工位模板")) {
			TableInfo[] infos = new TableInfo[25];
			infos[0] = new TableInfo("首页", "", false);
			infos[1] = new TableInfo("有效页", "", false);
			infos[2] = new TableInfo("构成表", "", false);
			infos[3] = new TableInfo("构成图", "", false);
			infos[4] = new TableInfo("式样差", "", false);
			infos[5] = new TableInfo("涂胶", "", false);
			infos[6] = new TableInfo("螺柱焊", "", false);
			infos[7] = new TableInfo("螺母焊", "", false);
			infos[8] = new TableInfo("螺栓焊", "", false);
			infos[9] = new TableInfo("点焊-PSW", "1", true);
			infos[10] = new TableInfo("点焊-RSW", "1", true);
			infos[11] = new TableInfo("点焊-MSW", "1", true);
			infos[12] = new TableInfo("点焊-SSW", "1", true);
			infos[13] = new TableInfo(" PSW", "", false);
			infos[14] = new TableInfo("RSW气动", "", false);
			infos[15] = new TableInfo("RSW伺服", "", false);
			infos[16] = new TableInfo("临时参数", "", false);
			infos[17] = new TableInfo("弧焊作业", "", false);
			infos[18] = new TableInfo("HEM指示", "", false);
			infos[19] = new TableInfo("铰链安装", "", false);
			infos[20] = new TableInfo("装配", "", false);
			infos[21] = new TableInfo("打点统计表", "", false);
			infos[22] = new TableInfo("NEPID安装", "", false);
			infos[23] = new TableInfo("其他", "", false);
			infos[24] = new TableInfo("法兰边检查", "", false);
			tableViewer.setInput(infos);
		} else if(model.equals("VIN码打刻模板")) {
			TableInfo[] infos = new TableInfo[5];
			infos[0] = new TableInfo("首页", "", false);
			infos[1] = new TableInfo("有效页", "", false);
			infos[2] = new TableInfo("构成表", "", false);
			infos[3] = new TableInfo("构成图", "", false);
			infos[4] = new TableInfo("打刻作业", "", false);
			tableViewer.setInput(infos);
		}else {
			TableInfo[] infos = new TableInfo[14];
			infos[0] = new TableInfo("首页", "", false);
			infos[1] = new TableInfo("有效页", "", false);
			infos[2] = new TableInfo("构成表", "", false);
			infos[3] = new TableInfo("构成图", "", false);
			infos[4] = new TableInfo("毛刺打磨", "", false);
			infos[5] = new TableInfo("法兰边检查", "", false);
			infos[6] = new TableInfo("拉铆", "", false);
			infos[7] = new TableInfo("装配", "", false);
			infos[8] = new TableInfo("安装", "", false);
			infos[9] = new TableInfo("外观检查", "", false);
			infos[10] = new TableInfo("电泳件", "", false);
			infos[11] = new TableInfo("电泳治具", "", false);
			infos[12] = new TableInfo("建付规格", "", false);
			infos[13] = new TableInfo("NEPID安装", "", false);
			tableViewer.setInput(infos);
		}
	}

	public static void main(String[] args) {
		SelectSheetTypeDialog dialog1 = new SelectSheetTypeDialog(new Shell(), SWT.SHELL_TRIM,null);
		dialog1.open();
	}
}
