package com.dfl.report.defects;

import java.awt.BorderLayout;
import java.awt.Dimension;
import java.awt.EventQueue;
import java.awt.Frame;

import javax.swing.JFrame;
import javax.swing.JPanel;
import javax.swing.GroupLayout.Alignment;
import javax.swing.LayoutStyle.ComponentPlacement;
import javax.swing.border.EmptyBorder;
import javax.swing.event.ChangeEvent;
import javax.swing.event.ChangeListener;

import com.dfl.report.dialog.SelectionStageDialog;
import com.dfl.report.handlers.AntirustRequirementsCheckOp;
import com.dfl.report.util.Util;
import com.teamcenter.rac.aif.AbstractAIFUIApplication;
import com.teamcenter.rac.aif.kernel.InterfaceAIFComponent;
import com.teamcenter.rac.aifrcp.AIFUtility;
import com.teamcenter.rac.kernel.TCComponent;
import com.teamcenter.rac.kernel.TCComponentBOMLine;
import com.teamcenter.rac.kernel.TCComponentForm;
import com.teamcenter.rac.kernel.TCComponentItem;
import com.teamcenter.rac.kernel.TCComponentItemRevision;
import com.teamcenter.rac.kernel.TCException;
import com.teamcenter.rac.kernel.TCPreferenceService;
import com.teamcenter.rac.kernel.TCSession;
import com.teamcenter.rac.testmanager.ui.model.TestManagerModelObject;

import javax.swing.GroupLayout;
import javax.swing.JButton;
import javax.swing.JLabel;
import javax.swing.JCheckBox;
import javax.swing.JComboBox;

import java.awt.GridBagLayout;
import java.awt.GridBagConstraints;
import java.awt.Insets;
import java.awt.SystemColor;
import java.awt.Toolkit;
import java.awt.event.ActionListener;
import java.util.ArrayList;
import java.util.List;
import java.awt.event.ActionEvent;

public class DefectsListDialog extends JFrame {

	private static final long serialVersionUID = 1L;
	private JPanel contentPane;
	private JComboBox comboBox;
	private AbstractAIFUIApplication app;
	private InterfaceAIFComponent[] aifComponents;
	private TCSession session;
	private  JCheckBox chckbxNewCheckBox_1;
	private  JCheckBox chckbxNewCheckBox_2;
	private  JCheckBox chckbxNewCheckBox_3;
	private  JCheckBox chckbxNewCheckBox_4;
	private  JCheckBox chckbxNewCheckBox_5;
	private  JCheckBox chckbxNewCheckBox;
	private ArrayList numlist = new ArrayList();
	private TCComponent folder;
	private String stage;
	
	public DefectsListDialog(AbstractAIFUIApplication app, TCComponent folder, InterfaceAIFComponent[] aifComponents, TCSession session) {
		setIconImage(Toolkit.getDefaultToolkit().getImage(SelectionStageDialog.class.getResource("/com/dfl/report/dialog/defaultapplication_16.png")));

		this.app = app;
		this.folder = folder;
		this.aifComponents = aifComponents;
		this.session = session;
		setTitle("\u9009\u62E9\u9636\u6BB5\u548C\u6D4B\u8BD5\u7ED3\u679C\u7C7B\u578B");
		setDefaultCloseOperation(JFrame.DISPOSE_ON_CLOSE);
		setBounds(100, 100, 455, 176);
		contentPane = new JPanel();
		contentPane.setBorder(new EmptyBorder(5, 5, 5, 5));
		contentPane.setLayout(new BorderLayout(0, 0));
		setContentPane(contentPane);

		JPanel panel = new JPanel();
		panel.setBackground(SystemColor.menu);
		contentPane.add(panel, BorderLayout.CENTER);

		JLabel label = new JLabel("\u9636\u6BB5\uFF1A");

		comboBox = new JComboBox();
		JLabel label_1 = new JLabel("\u6D4B\u8BD5\u7ED3\u679C\u7C7B\u578B:");
		
		chckbxNewCheckBox = new JCheckBox("1  [NG]");
		chckbxNewCheckBox.setBackground(SystemColor.menu);
		chckbxNewCheckBox.setSelected(true);
		chckbxNewCheckBox.setEnabled(false);
		panel.add(chckbxNewCheckBox);
		chckbxNewCheckBox_1 = new JCheckBox("2  [NG (Agreed)]");
		chckbxNewCheckBox_1.setBackground(SystemColor.menu);
		panel.add(chckbxNewCheckBox_1);
		chckbxNewCheckBox_2 = new JCheckBox("3  [No data]");
		chckbxNewCheckBox_2.setBackground(SystemColor.menu);
		panel.add(chckbxNewCheckBox_2);
		chckbxNewCheckBox_3 = new JCheckBox("4  [Item N/A]");
		chckbxNewCheckBox_3.setBackground(SystemColor.menu);
		panel.add(chckbxNewCheckBox_3);
		chckbxNewCheckBox_4 = new JCheckBox("5  [Inapplicability]");
		chckbxNewCheckBox_4.setBackground(SystemColor.menu);
		panel.add(chckbxNewCheckBox_4);
		chckbxNewCheckBox_5 = new JCheckBox("9  [Other]");
		chckbxNewCheckBox_5.setBackground(SystemColor.menu);
		panel.add(chckbxNewCheckBox_5);
		//ArrayList list = getStages(app);
		ArrayList list = getSelectStateRule();
		for (int i = 0; i < list.size(); i++) {
			comboBox.addItem(list.get(i).toString());
		}
		// 确定事件
		JButton button = new JButton("\u786E\u5B9A");
		button.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent arg0) {	
				getCheckBoxValue();
				int index = comboBox.getSelectedIndex();
				if(index==-1) {
					stage = null;
				}else {
					stage = comboBox.getSelectedItem().toString();
				}							
				DefectsListDialog.this.dispose();
				
				Thread thread = new Thread() {
					public void run() {
						CreateOp(stage,numlist);		
					}
				};
				thread.start();				
			}
		});

		// 取消事件
		JButton button_1 = new JButton("\u53D6\u6D88");
		button_1.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent arg0) {
				DefectsListDialog.this.dispose();
			}
		});
		
		
		GroupLayout gl_panel = new GroupLayout(panel);
		gl_panel.setHorizontalGroup(
			gl_panel.createParallelGroup(Alignment.LEADING)
				.addGroup(gl_panel.createSequentialGroup()
					.addGroup(gl_panel.createParallelGroup(Alignment.LEADING)
						.addGroup(gl_panel.createSequentialGroup()
							.addGap(20)
							.addGroup(gl_panel.createParallelGroup(Alignment.TRAILING)
								.addGroup(Alignment.LEADING, gl_panel.createSequentialGroup()
									.addComponent(label)
									.addPreferredGap(ComponentPlacement.RELATED)
									.addComponent(comboBox, GroupLayout.PREFERRED_SIZE, 163, GroupLayout.PREFERRED_SIZE))
								.addGroup(gl_panel.createSequentialGroup()
									.addGroup(gl_panel.createParallelGroup(Alignment.LEADING)
										.addComponent(chckbxNewCheckBox_3)
										.addComponent(chckbxNewCheckBox)
										.addComponent(label_1))
									.addGap(8)
									.addGroup(gl_panel.createParallelGroup(Alignment.LEADING)
										.addComponent(chckbxNewCheckBox_1)
										.addComponent(chckbxNewCheckBox_4))))
							.addGap(15)
							.addGroup(gl_panel.createParallelGroup(Alignment.LEADING)
								.addComponent(chckbxNewCheckBox_5)
								.addComponent(chckbxNewCheckBox_2)))
						.addGroup(gl_panel.createSequentialGroup()
							.addGap(96)
							.addComponent(button, GroupLayout.PREFERRED_SIZE, 73, GroupLayout.PREFERRED_SIZE)
							.addGap(60)
							.addComponent(button_1, GroupLayout.PREFERRED_SIZE, 69, GroupLayout.PREFERRED_SIZE)))
					.addContainerGap(36, Short.MAX_VALUE))
		);
		gl_panel.setVerticalGroup(
			gl_panel.createParallelGroup(Alignment.TRAILING)
				.addGroup(gl_panel.createSequentialGroup()
					.addGap(22)
					.addGroup(gl_panel.createParallelGroup(Alignment.BASELINE)
						.addComponent(label)
						.addComponent(comboBox, GroupLayout.PREFERRED_SIZE, GroupLayout.DEFAULT_SIZE, GroupLayout.PREFERRED_SIZE))
					.addGap(18)
					.addComponent(label_1)
					.addPreferredGap(ComponentPlacement.UNRELATED)
					.addGroup(gl_panel.createParallelGroup(Alignment.BASELINE)
						.addComponent(chckbxNewCheckBox)
						.addComponent(chckbxNewCheckBox_2)
						.addComponent(chckbxNewCheckBox_1))
					.addGap(18)
					.addGroup(gl_panel.createParallelGroup(Alignment.BASELINE)
						.addComponent(chckbxNewCheckBox_3)
						.addComponent(chckbxNewCheckBox_5)
						.addComponent(chckbxNewCheckBox_4))
					.addPreferredGap(ComponentPlacement.RELATED, 22, Short.MAX_VALUE)
					.addGroup(gl_panel.createParallelGroup(Alignment.BASELINE)
						.addComponent(button_1)
						.addComponent(button))
					.addContainerGap())
		);
		panel.setLayout(gl_panel);

		// 居中显示
		int width = 480;
		int height = 300;
		Dimension dimension1 = Toolkit.getDefaultToolkit().getScreenSize();
		int i = (dimension1.width - width) / 2;
		int j = (dimension1.height - height) / 2;
		setBounds(i, j, width, height);
	}
	private ArrayList getStages(AbstractAIFUIApplication app) {
		// TODO Auto-generated method stub
		ArrayList stages = new ArrayList();// 获取当前BOP所有关联的测试用例的测试阶段
		InterfaceAIFComponent[] targets = app.getTargetComponents();
		TCComponentBOMLine targetbl = (TCComponentBOMLine) targets[0];
		TCComponentBOMLine topbl = null;
		try {
			topbl = targetbl.window().getTopBOMLine();
		} catch (TCException e1) {
			// TODO Auto-generated catch block
			e1.printStackTrace();
		}
		// 获取所有焊装工位关联的测试用例
		try {
			ArrayList list = Util.SearchTests(topbl);
			if(list!=null&&list.size()>0) {
				for (int i = 0; i < list.size(); i++) {
					TestManagerModelObject modelObject = (TestManagerModelObject) list.get(i);
					// 获取测试用例实例
					TCComponent testCaseInstance = modelObject.getTestComponent();

					// 获取测试用例
					TCComponent testCase = modelObject.getTestCase();

					// 根据要件类型取值
					String testCaseType = Util.getRelProperty(testCase, "b8_TestCaseType");

					// 取生产要件
					if (testCaseType.equals("0")) {
						if ((testCase != null && testCase instanceof TCComponentItem)
								&& (testCaseInstance != null && testCaseInstance instanceof TCComponentForm)) {
							// 获取测试用例最新版本
							TCComponentItem testCaseItem = (TCComponentItem) testCase;
							TCComponentItemRevision testCaseRev = testCaseItem.getLatestItemRevision();
							// 获取测试实例活动对象表单（测试活动）
							TCComponent[] activitys = testCaseInstance.getRelatedComponents("Tm0TestInstanceActivityRel");

							if (activitys != null && activitys.length > 0) {

								List tempList = new ArrayList();
								for (int j = 0; j < activitys.length; j++) {
									tempList.add(activitys[j]);
								}

								for (int j = 0; j < tempList.size(); j++) {
									TCComponentForm testactivity = (TCComponentForm) tempList.get(j);
									String str = Util.getProperty(testactivity, "b8_TestStage");// 阶段
									if (!stages.contains(str)) {
										stages.add(str);
									}
								}
							}

						}
					}
				}
			}
			
			return stages;

		} catch (Exception e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		return stages;
	}
	// 查询阶段首选项，获取阶段信息
	private ArrayList getSelectStateRule() {
		ArrayList rule = new ArrayList();
		try {
			AbstractAIFUIApplication app = AIFUtility.getCurrentApplication();
			TCSession session = (TCSession) app.getSession();
			TCPreferenceService preferenceService = session.getPreferenceService();

			String str = preferenceService.getPreferenceDescription("DFL9_Selection_test_phase");
			if (str != null) {
				String[] values = preferenceService.getStringValues("DFL9_Selection_test_phase");
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

	private void CreateOp(String stage, ArrayList numlist) {
		// TODO Auto-generated method stub
		DefectsListOp op = new DefectsListOp(app, stage,numlist,folder,aifComponents,session);
	}

	private void getCheckBoxValue() {
		// TODO Auto-generated method stub
		
		String val = chckbxNewCheckBox.getText();
		numlist.add(val);
		if (chckbxNewCheckBox_1.isSelected()) {
			String val2 = chckbxNewCheckBox_1.getText();
			if (!numlist.contains(val2)) {
				numlist.add(val2);
			}
		}
		if (chckbxNewCheckBox_2.isSelected()) {
			String val3 = chckbxNewCheckBox_2.getText();
			if (!numlist.contains(val3)) {
				numlist.add(val3);
			}
		}
		if (chckbxNewCheckBox_3.isSelected()) {
			String val4 = chckbxNewCheckBox_3.getText();
			if (!numlist.contains(val4)) {
				numlist.add(val4);
			}
		}
		if (chckbxNewCheckBox_4.isSelected()) {
			String val5 = chckbxNewCheckBox_4.getText();
			if (!numlist.contains(val5)) {
				numlist.add(val5);
			}
		}
		if (chckbxNewCheckBox_5.isSelected()) {
			String val6 = chckbxNewCheckBox_5.getText();
			if (!numlist.contains(val6)) {
				numlist.add(val6);
			}
		}
	}

}
