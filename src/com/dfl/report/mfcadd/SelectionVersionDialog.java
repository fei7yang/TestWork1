package com.dfl.report.mfcadd;

import java.awt.BorderLayout;
import java.awt.Dimension;
import java.awt.EventQueue;
import java.awt.Toolkit;

import javax.swing.JFrame;
import javax.swing.JPanel;
import javax.swing.border.EmptyBorder;

import com.dfl.report.handlers.AntirustRequirementsCheckOp;
import com.dfl.report.util.Util;
import com.teamcenter.rac.aif.AbstractAIFUIApplication;
import com.teamcenter.rac.aif.kernel.InterfaceAIFComponent;
import com.teamcenter.rac.kernel.TCComponent;
import com.teamcenter.rac.kernel.TCComponentBOMLine;
import com.teamcenter.rac.kernel.TCComponentForm;
import com.teamcenter.rac.kernel.TCComponentItem;
import com.teamcenter.rac.kernel.TCComponentItemRevision;
import com.teamcenter.rac.kernel.TCException;
import com.teamcenter.rac.testmanager.ui.model.TestManagerModelObject;

import javax.swing.GroupLayout;
import javax.swing.GroupLayout.Alignment;
import javax.swing.JLabel;
import javax.swing.JTextField;
import javax.swing.LayoutStyle.ComponentPlacement;
import javax.swing.JComboBox;
import javax.swing.JButton;
import java.awt.event.ActionListener;
import java.util.ArrayList;
import java.util.Collections;
import java.util.Comparator;
import java.util.List;
import java.awt.event.ActionEvent;
import java.awt.Color;
import java.awt.SystemColor;

public class SelectionVersionDialog extends JFrame {

	private JPanel contentPane;
	private JComboBox comboBox;
	private AbstractAIFUIApplication app;
	private TCComponent folder;
	private String stage;

	/**
	 * Launch the application.
	 */
//	public static void main(String[] args) {
//		EventQueue.invokeLater(new Runnable() {
//			public void run() {
//				try {
//					SelectionStageDialog frame = new SelectionStageDialog();
//					frame.setVisible(true);
//				} catch (Exception e) {
//					e.printStackTrace();
//				}
//			}
//		});
//	}

	/**
	 * Create the frame.
	 * 
	 * @param app
	 * @param folder 
	 */
	public SelectionVersionDialog(AbstractAIFUIApplication app, TCComponent folder) {
		setIconImage(Toolkit.getDefaultToolkit().getImage(SelectionVersionDialog.class.getResource("/com/dfl/report/dialog/defaultapplication_16.png")));

		this.app = app;
		this.folder = folder;

		setTitle("\u9009\u62E9\u7248\u6B21");
		setDefaultCloseOperation(JFrame.DISPOSE_ON_CLOSE);
		setBounds(100, 100, 455, 176);
		contentPane = new JPanel();
		contentPane.setBorder(new EmptyBorder(5, 5, 5, 5));
		contentPane.setLayout(new BorderLayout(0, 0));
		setContentPane(contentPane);

		JPanel panel = new JPanel();
		panel.setBackground(SystemColor.window);
		contentPane.add(panel, BorderLayout.CENTER);

		JLabel label = new JLabel("\u7248\u6B21\uFF1A");

		comboBox = new JComboBox();

		ArrayList list = getStages(app);
		for (int i = 0; i < list.size(); i++) {
			comboBox.addItem(list.get(i).toString());
		}
		// 确定事件
		JButton button = new JButton("\u786E\u5B9A");
		button.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent arg0) {			
				int index = comboBox.getSelectedIndex();				
				if(index==-1) {
					stage = null;
				}else {
					stage = comboBox.getSelectedItem().toString();
				}	
				SelectionVersionDialog.this.dispose();
				Thread thread = new Thread() {
					public void run() {
						CreateOp(stage);			
					}
				};
				thread.start();
								
			}
		});

		// 取消事件
		JButton button_1 = new JButton("\u53D6\u6D88");
		button_1.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent arg0) {
				SelectionVersionDialog.this.dispose();
			}
		});
		GroupLayout gl_panel = new GroupLayout(panel);
		gl_panel.setHorizontalGroup(
				gl_panel.createParallelGroup(Alignment.TRAILING)
						.addGroup(Alignment.LEADING,
								gl_panel.createSequentialGroup().addGap(20).addComponent(label)
										.addPreferredGap(ComponentPlacement.RELATED)
										.addComponent(comboBox, GroupLayout.PREFERRED_SIZE, 163,
												GroupLayout.PREFERRED_SIZE)
										.addContainerGap(206, Short.MAX_VALUE))
						.addGroup(gl_panel.createSequentialGroup().addContainerGap(267, Short.MAX_VALUE)
								.addComponent(button, GroupLayout.PREFERRED_SIZE, 73, GroupLayout.PREFERRED_SIZE)
								.addPreferredGap(ComponentPlacement.UNRELATED)
								.addComponent(button_1, GroupLayout.PREFERRED_SIZE, 69, GroupLayout.PREFERRED_SIZE)
								.addContainerGap()));
		gl_panel.setVerticalGroup(gl_panel.createParallelGroup(Alignment.LEADING).addGroup(gl_panel
				.createSequentialGroup().addGap(22)
				.addGroup(gl_panel.createParallelGroup(Alignment.BASELINE).addComponent(label).addComponent(comboBox,
						GroupLayout.PREFERRED_SIZE, GroupLayout.DEFAULT_SIZE, GroupLayout.PREFERRED_SIZE))
				.addContainerGap(84, Short.MAX_VALUE))
				.addGroup(Alignment.TRAILING,
						gl_panel.createSequentialGroup().addContainerGap(69, Short.MAX_VALUE).addGroup(gl_panel
								.createParallelGroup(Alignment.BASELINE).addComponent(button).addComponent(button_1))
								.addGap(35)));
		panel.setLayout(gl_panel);

		// 居中显示
		int width = 400;
		int height = 250;
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

					// 取要件类型为防锈的
					if (testCaseType.equals("2")) {
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

	private void CreateOp(String stage) {
		// TODO Auto-generated method stub
		//AntirustRequirementsCheckOp op = new AntirustRequirementsCheckOp(app, stage,folder);
	}
}
