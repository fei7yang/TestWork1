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

public class SelectTemplate extends Dialog {

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
	protected Shell shell;
//	private Table table;
	private TableViewer tableViewer;
	private Combo combo_1;
	public LinkedHashMap<String, String> map = new LinkedHashMap<String, String>();
	public ArrayList list = new ArrayList();
	public String nameNO;
	public String model;
	public boolean IsSameout;
	private Label lblNewLabel;
	private Text text;
	private Combo combo;

	/**
	 * Create the dialog.
	 * 
	 * @param parent
	 * @param style
	 */
	public SelectTemplate(Shell parent, int style) {
		super(parent, style);
		setText("SWT Dialog");
	}

	/**
	 * Open the dialog.
	 * 
	 * @return the result
	 */
	public Object open() {
		createContents();
		shell.open();
		shell.layout();
		Display display = getParent().getDisplay();
		centerToScreen(display);
		while (!shell.isDisposed()) {
			if (!display.readAndDispatch()) {
				display.sleep();
			}
		}
		return result;
	}

	protected void centerToScreen(Display display) {
		int nLocationX = display.getClientArea().width / 2 - shell.getSize().x / 2;
		int nLocationY = display.getClientArea().height / 2 - shell.getSize().y / 2;
		shell.setLocation(nLocationX, nLocationY);
	}

	/**
	 * Create contents of the dialog.
	 */
	private void createContents() {
		shell = new Shell(getParent(), getStyle());
		shell.setMinimumSize(new Point(136, 20));
		//shell.setImage(SWTResourceManager.getImage(SelectTemplate.class,"/com/dfl/report/workschedule/defaultapplication_16.png"));
		shell.setImage(SWTResourceManager.getImage(SelectTemplate.class,"/com/dfl/report/imags/defaultapplication_16.png"));
		shell.setSize(600, 621);
		shell.setText("\u9009\u62E9\u6A21\u677F");
		shell.setLayout(new BorderLayout(0, 0));

		Composite composite = new Composite(shell, SWT.NONE);
		composite.setLayoutData(BorderLayout.NORTH);
		composite.setLayout(null);

		Label label = new Label(composite, SWT.NONE);
		label.setBounds(211, 9, 67, 17);
		label.setText("\u547D\u540D\u5E8F\u53F7\uFF1A");

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
		combo_1.setBounds(72, 6, 120, 20);
		combo_1.add("普通工位模板");
		combo_1.add("VIN码打刻模板");
		combo_1.add("调整线模板");
		combo_1.select(0);
		
		text = new Text(composite, SWT.BORDER);
		text.setBounds(284, 6, 88, 23);
		
		Label label_2 = new Label(composite, SWT.NONE);
		label_2.setBounds(389, 9, 88, 17);
		label_2.setText("\u5DE6\u53F3\u5DE5\u4F4D\u540C\u51FA\uFF1A");
		
		combo = new Combo(composite, SWT.NONE);
		combo.setBounds(483, 6, 75, 25);
		combo.add("是");
		combo.add("否");
		combo.select(0);

		Composite composite_1 = new Composite(shell, SWT.NONE);
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
				shell.dispose();
			}
		});
		btnNewButton_1.setText("\u53D6\u6D88");
		
		lblNewLabel = new Label(composite_1, SWT.NONE);
		lblNewLabel.setForeground(SWTResourceManager.getColor(SWT.COLOR_RED));
		lblNewLabel.setBounds(10, 10, 372, 17);

		Composite composite_2 = new Composite(shell, SWT.NONE);
		composite_2.setLayoutData(BorderLayout.EAST);
		composite_2.setLayout(new RowLayout(SWT.VERTICAL));

		Button moveUpButton = new Button(composite_2, SWT.NONE);
		moveUpButton.addSelectionListener(new SelectionAdapter() {
			@Override
			public void widgetSelected(SelectionEvent e) {
				moveUpAction();
			}
		});
		moveUpButton.setImage(
				SWTResourceManager.getImage(SelectTemplate.class, "/com/dfl/report/imags/moveup_16.png"));

		Button moveDownButton = new Button(composite_2, SWT.NONE);
		moveDownButton.addSelectionListener(new SelectionAdapter() {
			@Override
			public void widgetSelected(SelectionEvent e) {
				moveDownAction();
			}
		});
		moveDownButton.setImage(
				SWTResourceManager.getImage(SelectTemplate.class, "/com/dfl/report/imags/movedown_16.png"));

		// 全选事件
		Button button = new Button(composite_2, SWT.NONE);
		button.addSelectionListener(new SelectionAdapter() {
			@Override
			public void widgetSelected(SelectionEvent e) {
				AllSelectionOp();
			}
		});
		button.setText("\u5168\u90E8\u9009\u62E9");

		// 取消事件
		Button button_1 = new Button(composite_2, SWT.NONE);
		button_1.addSelectionListener(new SelectionAdapter() {
			@Override
			public void widgetSelected(SelectionEvent e) {
				AllCancelOp();
			}
		});
		button_1.setText("\u5168\u90E8\u53D6\u6D88");

		tableViewer = new TableViewer(shell, SWT.BORDER | SWT.CHECK | SWT.FULL_SELECTION | SWT.MULTI);
		Table table = tableViewer.getTable();
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
		tblclmnNewColumn.setWidth(222);
		tblclmnNewColumn.setText("Sheet\u540D\u79F0");

		TableColumn tblclmnPage = new TableColumn(table, SWT.NONE);
		tblclmnPage.setWidth(100);
		tblclmnPage.setText("\u9875\u6570");

		CellEditor acelleditor[] = new CellEditor[2];
		acelleditor[1] = new TextCellEditor(table, 131072);
		tableViewer.setCellEditors(acelleditor);
		
//		ArrayList list = getSizeRule();
//		if(list!=null && list.size()>0) {
//			for(int i=0;i<list.size();i++) {
//				combo.add((String) list.get(i));
//			}
//		}else {
//			lblNewLabel.setText("版次信息首选项(DFL9_get_version_information)未设置！");
//		}

		loadTableData();

	}

	protected void ComformEvent() {
		// TODO Auto-generated method stub
		if (combo_1.getSelectionIndex() == -1) {
			lblNewLabel.setText("请选择模板类型！");
			return;
		}
		if (text.getText().isEmpty()) {
			lblNewLabel.setText("请填写命名序号！");
			return;
		}
		if (combo.getText().isEmpty()) {
			lblNewLabel.setText("请选择左右工位是否同出！");
			return;
		}
		if(combo.getText().equals("是")) {
			IsSameout = true;
		}else {
			IsSameout = false;
		}		
		model = combo_1.getItem(combo_1.getSelectionIndex());
		nameNO = text.getText();
		Table table = tableViewer.getTable();
		TableItem[] items = table.getItems();// 获取表格所有的行
		String rswqd = "";
		String rswsf = "";
		for (int i = 0; i < items.length; i++) {
			TableItem tableItem = items[i];
			if (tableItem.getChecked()) {
				String[] str = new String[2];
				str[0] = tableItem.getText(0);
				str[1] = tableItem.getText(1);
				if(str[0].equals("RSW气动")) {
					rswqd = str[0];
				}
				if(str[0].equals("RSW伺服")) {
					rswsf = str[0];
				}
				if(str[1]==null || str[1].isEmpty()) {
					str[1] = "1";
				}
				map.put(str[0], str[1]);
				list.add(str[0]);
			}
		}
		if(map.size()<1) {
			lblNewLabel.setText("请勾选需要输出的sheet模板！");
			return;
		}
		if(!rswqd.isEmpty()&&!rswsf.isEmpty()) {
			lblNewLabel.setText("RSW气动和RSW伺服只能选一种！");
			rswqd = "";
			rswsf = "";
			list.clear();
			for(Iterator<java.util.Map.Entry<String, String>> it = map.entrySet().iterator();it.hasNext();) {
				java.util.Map.Entry<String, String> item = it.next();
				it.remove();
			}

			return;
		}
		
		shell.dispose();
	}

	// 全选操作
	protected void AllSelectionOp() {
		Table table = tableViewer.getTable();
		TableItem[] items = table.getItems();// 获取表格所有的行
		ArrayList list = new ArrayList();
		for (int i = 0; i < items.length; i++) {
			TableItem tableItem = items[i];
			tableItem.setChecked(true);
			// tableItem.getChecked();
		}
	}

	// 全部取消操作
	protected void AllCancelOp() {
		Table table = tableViewer.getTable();
		TableItem[] items = table.getItems();// 获取表格所有的行
		ArrayList list = new ArrayList();
		for (int i = 0; i < items.length; i++) {
			TableItem tableItem = items[i];
			tableItem.setChecked(false);
			// tableItem.getChecked();
		}
	}

	protected void moveDownAction() {
		// TODO Auto-generated method stub
		Table table = tableViewer.getTable();

		int[] selectIndices = table.getSelectionIndices();

		if (selectIndices == null || selectIndices.length <= 0) {
			return;
		}

		int[] selectIndices2 = new int[selectIndices.length];

		// 选择的不是连续的项,不处理
		{
			int tempv = 0;
			for (int i = 0; i < selectIndices.length; i++) {

				selectIndices2[i] = selectIndices[i] + 1;
				if (i == 0) {
					tempv = selectIndices[i];
				} else {
					int t = selectIndices[i] - tempv;
					if (t != 1) {
						System.out.println("选择的不是连续的项，不处理");
						return;
					}
					tempv = selectIndices[i];
				}

			}
		}

		int startRow = selectIndices[0];
		int maxRow = selectIndices[selectIndices.length - 1];

		// 选择的项已经在最后,不处理
		{
			if (maxRow == table.getItemCount() - 1) {
				System.out.println("选择的项已经在最后,不处理");
				return;
			}
		}

		// 将选择的行向下移动
		{

			TableItem[] items = table.getItems();// 获取表格所有的行
			ArrayList list = new ArrayList();
			java.util.Map<TableInfo,Boolean> map = new HashMap<>();
			for (int i = 0; i < items.length; i++) {
				TableItem tableItem = items[i];
				// tableItem.getChecked();				
				TableInfo info = (TableInfo) tableItem.getData();
				map.put(info, tableItem.getChecked());
				list.add(info);
			}
			System.out.println("startRow:" + startRow);
			System.out.println("maxRow:" + maxRow);

			ArrayList infoList = MoveUtil.moveDown(list, startRow, maxRow);

			for (int i = 0; i < infoList.size(); i++) {
				TableInfo info = (TableInfo) infoList.get(i);
				System.out.println("info:" + info.getName());
			}

			TableInfo[] newInfos = (TableInfo[]) infoList.toArray(new TableInfo[infoList.size()]);

			table.removeAll();
			tableViewer.setInput(newInfos);
			
			TableItem[] item2 = table.getItems();// 获取表格所有的行
			for (int i = 0; i < item2.length; i++) {
				if(map.containsKey(item2[i].getData()))
				{
					item2[i].setChecked(map.get(item2[i].getData()));
				}
			}

			// 选中移动后的行
			table.setSelection(selectIndices2);

		}

	}

	protected void moveUpAction() {
		// TODO Auto-generated method stub
		Table table = tableViewer.getTable();

		int[] selectIndices = table.getSelectionIndices();

		if (selectIndices == null || selectIndices.length <= 0) {
			return;
		}

		int[] selectIndices2 = new int[selectIndices.length];

		// 选择的不是连续的项,不处理
		{
			int tempv = 0;
			for (int i = 0; i < selectIndices.length; i++) {

				selectIndices2[i] = selectIndices[i] - 1;
				if (i == 0) {
					tempv = selectIndices[i];
				} else {
					int t = selectIndices[i] - tempv;
					if (t != 1) {
						System.out.println("选择的不是连续的项，不处理");
						return;
					}
					tempv = selectIndices[i];
				}

			}
		}

		int startRow = selectIndices[0];
		int maxRow = selectIndices[selectIndices.length - 1];

		// 选择的项已经是第一项,不处理
		{
			if (maxRow == 0) {
				System.out.println("选择的项已经是第一项,不处理");
				return;
			}
		}

		// 将选择的行向上移动
		{

			TableItem[] items = table.getItems();// 获取表格所有的行
			ArrayList list = new ArrayList();
			java.util.Map<TableInfo,Boolean> map = new HashMap<>();
			for (int i = 0; i < items.length; i++) {
				TableItem tableItem = items[i];
				// tableItem.getChecked();
				TableInfo info = (TableInfo) tableItem.getData();
				map.put(info, tableItem.getChecked());
				list.add(info);
			}
			System.out.println("startRow:" + startRow);
			System.out.println("maxRow:" + maxRow);

			ArrayList infoList = MoveUtil.moveUp(list, startRow, maxRow);

			for (int i = 0; i < infoList.size(); i++) {
				TableInfo info = (TableInfo) infoList.get(i);
				System.out.println("info:" + info.getName());
			}

			TableInfo[] newInfos = (TableInfo[]) infoList.toArray(new TableInfo[infoList.size()]);

			table.removeAll();
			tableViewer.setInput(newInfos);
			
			TableItem[] item2 = table.getItems();// 获取表格所有的行
			for (int i = 0; i < item2.length; i++) {
				if(map.containsKey(item2[i].getData()))
				{
					item2[i].setChecked(map.get(item2[i].getData()));
				}
			}

			// 选中移动后的行
			table.setSelection(selectIndices2);

		}

	}

	private void loadTableData() {
		// TODO Auto-generated method stub
		TableInfo[] infos = new TableInfo[25];
		infos[0] = new TableInfo("首页", "", false);
		infos[1] = new TableInfo("有效页", "", false);
		infos[2] = new TableInfo("构成表", "", false);
		infos[3] = new TableInfo("构成图", "", false);
		infos[4] = new TableInfo("式样差", "", false);
		infos[5] = new TableInfo("涂胶", "1", true);
		infos[6] = new TableInfo("螺柱焊", "1", true);
		infos[7] = new TableInfo("螺母焊", "1", true);
		infos[8] = new TableInfo("螺栓焊", "1", true);
		infos[9] = new TableInfo("点焊-PSW", "1", true);
		infos[10] = new TableInfo("点焊-RSW", "1", true);
		infos[11] = new TableInfo("点焊-MSW", "1", true);
		infos[12] = new TableInfo("点焊-SSW", "1", true);
		infos[13] = new TableInfo(" PSW", "", false);
		infos[14] = new TableInfo("RSW气动", "", false);
		infos[15] = new TableInfo("RSW伺服", "", false);
		infos[16] = new TableInfo("临时参数", "1", true);
		infos[17] = new TableInfo("弧焊作业", "1", true);
		infos[18] = new TableInfo("HEM指示", "1", true);
		infos[19] = new TableInfo("铰链安装", "1", true);
		infos[20] = new TableInfo("装配", "1", true);
		infos[21] = new TableInfo("打点统计表", "", false);
		infos[22] = new TableInfo("NEPID安装", "1", true);
		infos[23] = new TableInfo("其他", "1", true);
		infos[24] = new TableInfo("法兰边检查", "1", true);
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
			infos[5] = new TableInfo("涂胶", "1", true);
			infos[6] = new TableInfo("螺柱焊", "1", true);
			infos[7] = new TableInfo("螺母焊", "1", true);
			infos[8] = new TableInfo("螺栓焊", "1", true);
			infos[9] = new TableInfo("点焊-PSW", "1", true);
			infos[10] = new TableInfo("点焊-RSW", "1", true);
			infos[11] = new TableInfo("点焊-MSW", "1", true);
			infos[12] = new TableInfo("点焊-SSW", "1", true);
			infos[13] = new TableInfo(" PSW", "", false);
			infos[14] = new TableInfo("RSW气动", "", false);
			infos[15] = new TableInfo("RSW伺服", "", false);
			infos[16] = new TableInfo("临时参数", "1", true);
			infos[17] = new TableInfo("弧焊作业", "1", true);
			infos[18] = new TableInfo("HEM指示", "1", true);
			infos[19] = new TableInfo("铰链安装", "1", true);
			infos[20] = new TableInfo("装配", "1", true);
			infos[21] = new TableInfo("打点统计表", "", false);
			infos[22] = new TableInfo("NEPID安装", "1", true);
			infos[23] = new TableInfo("其他", "1", true);
			infos[24] = new TableInfo("法兰边检查", "1", true);
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

	// 查询版次首选项，获取版次信息
	private ArrayList getSizeRule() {
		ArrayList rule = new ArrayList();
		try {
			AbstractAIFUIApplication app = AIFUtility.getCurrentApplication();
			TCSession session = (TCSession) app.getSession();
			TCPreferenceService preferenceService = session.getPreferenceService();

			String str = preferenceService.getPreferenceDescription("DFL9_get_version_information");
			if (str != null) {
				String[] values = preferenceService.getStringValues("DFL9_get_version_information");
				for (int i = 0; i < values.length; i++) {
					String value = values[i];
					rule.add(value);
				}
			}
			return rule;
		} catch (Exception e) {
			e.printStackTrace();
		}
		return rule;
	}

	public static void main(String[] args) {
		SelectTemplate dialog1 = new SelectTemplate(new Shell(), SWT.SHELL_TRIM);
		dialog1.open();
	}
}
