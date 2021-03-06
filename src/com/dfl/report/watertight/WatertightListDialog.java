package com.dfl.report.watertight;

import java.awt.BorderLayout;
import java.awt.Color;
import java.awt.Dimension;
import java.awt.SystemColor;
import java.awt.Toolkit;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.List;

import javax.swing.GroupLayout;
import javax.swing.JButton;
import javax.swing.JComboBox;
import javax.swing.JFrame;
import javax.swing.JLabel;
import javax.swing.JPanel;
import javax.swing.GroupLayout.Alignment;
import javax.swing.LayoutStyle.ComponentPlacement;
import javax.swing.border.EmptyBorder;

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

public class WatertightListDialog extends JFrame{
	private JPanel contentPane;
	private JComboBox comboBox;
	private AbstractAIFUIApplication app;
	private TCComponent folder;
	private String stage;
	private InterfaceAIFComponent[] aifComponents;
	private TCSession session;
	private ArrayList rule;
	private JLabel lblDs;
	/**
	 * Launch the application.
	 */
//	public static void main(String[] args) {
//		EventQueue.invokeLater(new Runnable() {
//			public void run() {
//				try {
//					WatertightListDialog frame = new WatertightListDialog();
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
	 * @param session 
	 * @param aifComponents 
	 * @param rule 
	 * @param inputStream 
	 * @param folder2 
	 */
	public WatertightListDialog(AbstractAIFUIApplication app,  TCComponent folder, InterfaceAIFComponent[] aifComponents, TCSession session, ArrayList rule) {
		this.app = app;
		this.folder = folder;
		this.aifComponents =aifComponents;
		this.session = session;
		this.rule = rule;

		setIconImage(Toolkit.getDefaultToolkit().getImage(WatertightListDialog.class.getResource("/com/dfl/report/dialog/defaultapplication_16.png")));
		setTitle("\u9009\u62E9\u9636\u6BB5");
		setDefaultCloseOperation(JFrame.DISPOSE_ON_CLOSE);
		setBounds(100, 100, 455, 176);
		contentPane = new JPanel();
		contentPane.setBorder(new EmptyBorder(5, 5, 5, 5));
		contentPane.setLayout(new BorderLayout(0, 0));
		setContentPane(contentPane);

		JPanel panel = new JPanel();
		panel.setBackground(SystemColor.window);
		contentPane.add(panel, BorderLayout.CENTER);

		JLabel label = new JLabel("\u9636\u6BB5\uFF1A");

		comboBox = new JComboBox();

		//ArrayList list = getStages(app);
		//ArrayList list = getSelectStateRule();
		
		for (int i = 0; i < rule.size(); i++) {
			comboBox.addItem(rule.get(i).toString());
		}

		// ????????
		JButton button = new JButton("\u786E\u5B9A");
		button.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent arg0) {
				CreateOp();
			}
		});

		// ????????
		JButton button_1 = new JButton("\u53D6\u6D88");
		button_1.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent arg0) {
				WatertightListDialog.this.dispose();
			}
		});
		lblDs = new JLabel("");
		lblDs.setForeground(Color.RED);
		GroupLayout gl_panel = new GroupLayout(panel);
		gl_panel.setHorizontalGroup(
			gl_panel.createParallelGroup(Alignment.LEADING)
				.addGroup(gl_panel.createSequentialGroup()
					.addContainerGap(212, Short.MAX_VALUE)
					.addComponent(button, GroupLayout.PREFERRED_SIZE, 73, GroupLayout.PREFERRED_SIZE)
					.addPreferredGap(ComponentPlacement.UNRELATED)
					.addComponent(button_1, GroupLayout.PREFERRED_SIZE, 69, GroupLayout.PREFERRED_SIZE)
					.addContainerGap())
				.addGroup(gl_panel.createSequentialGroup()
					.addGap(20)
					.addGroup(gl_panel.createParallelGroup(Alignment.LEADING)
						.addComponent(lblDs, GroupLayout.PREFERRED_SIZE, 334, GroupLayout.PREFERRED_SIZE)
						.addGroup(gl_panel.createSequentialGroup()
							.addComponent(label)
							.addPreferredGap(ComponentPlacement.RELATED)
							.addComponent(comboBox, GroupLayout.PREFERRED_SIZE, 163, GroupLayout.PREFERRED_SIZE)))
					.addContainerGap(20, Short.MAX_VALUE))
		);
		gl_panel.setVerticalGroup(
			gl_panel.createParallelGroup(Alignment.TRAILING)
				.addGroup(gl_panel.createSequentialGroup()
					.addGap(22)
					.addGroup(gl_panel.createParallelGroup(Alignment.BASELINE)
						.addComponent(label)
						.addComponent(comboBox, GroupLayout.PREFERRED_SIZE, GroupLayout.DEFAULT_SIZE, GroupLayout.PREFERRED_SIZE))
					.addGap(27)
					.addComponent(lblDs, GroupLayout.PREFERRED_SIZE, 30, GroupLayout.PREFERRED_SIZE)
					.addPreferredGap(ComponentPlacement.RELATED, 43, Short.MAX_VALUE)
					.addGroup(gl_panel.createParallelGroup(Alignment.BASELINE)
						.addComponent(button)
						.addComponent(button_1))
					.addGap(35))
		);
		panel.setLayout(gl_panel);

		// ????????
		int width = 400;
		int height = 250;
		Dimension dimension1 = Toolkit.getDefaultToolkit().getScreenSize();
		int i = (dimension1.width - width) / 2;
		int j = (dimension1.height - height) / 2;
		setBounds(i, j, width, height);
	}

	private ArrayList getStages(AbstractAIFUIApplication app) {
		// TODO Auto-generated method stub
		ArrayList stages = new ArrayList();// ????????BOP????????????????????????????
		InterfaceAIFComponent[] targets = app.getTargetComponents();
		TCComponentBOMLine targetbl = (TCComponentBOMLine) targets[0];
		TCComponentBOMLine topbl = null;
		try {
			topbl = targetbl.window().getTopBOMLine();
		} catch (TCException e1) {
			// TODO Auto-generated catch block
			e1.printStackTrace();
		}
		// ??????????????????????????????
		try {
			ArrayList list = Util.SearchTests(topbl);
			if(list!=null&&list.size()>0) {
				for (int i = 0; i < list.size(); i++) {
					TestManagerModelObject modelObject = (TestManagerModelObject) list.get(i);
					// ????????????????
					TCComponent testCaseInstance = modelObject.getTestComponent();

					// ????????????
					TCComponent testCase = modelObject.getTestCase();

					// ????????????????
					String testCaseType = Util.getRelProperty(testCase, "b8_TestCaseType");

					// ??????????
					if (testCaseType.equals("1")) {
						if ((testCase != null && testCase instanceof TCComponentItem)
								&& (testCaseInstance != null && testCaseInstance instanceof TCComponentForm)) {
							// ????????????????????
							TCComponentItem testCaseItem = (TCComponentItem) testCase;
							TCComponentItemRevision testCaseRev = testCaseItem.getLatestItemRevision();
							// ????????????????????????????????????
							TCComponent[] activitys = testCaseInstance.getRelatedComponents("Tm0TestInstanceActivityRel");

							if (activitys != null && activitys.length > 0) {

								List tempList = new ArrayList();
								for (int j = 0; j < activitys.length; j++) {
									tempList.add(activitys[j]);
								}

								for (int j = 0; j < tempList.size(); j++) {
									TCComponentForm testactivity = (TCComponentForm) tempList.get(j);
									String str = Util.getProperty(testactivity, "b8_TestStage");// ????
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
		return null;
	}
	// ????????????????????????????
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
	private void CreateOp() {
		// TODO Auto-generated method stub
		int index = comboBox.getSelectedIndex();
		if (index == -1) {
			stage = null;
		} else {
			stage = comboBox.getSelectedItem().toString();
		}
		if(stage==null || stage.isEmpty()) {
			lblDs.setText("????????????");
		}
		WatertightListDialog.this.dispose();
		Thread thread = new Thread() {
			public void run() {
				WatertightListOp op = new WatertightListOp(app, stage,folder,aifComponents,session);
			}
		};
		thread.start();
		
	}

}
