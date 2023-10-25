package com.model;

import java.awt.Desktop;
import java.awt.GridBagConstraints;
import java.awt.GridBagLayout;
import java.awt.Insets;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.HashSet;
import java.util.Set;

import javax.swing.JButton;
import javax.swing.JComboBox;
import javax.swing.JFrame;
import javax.swing.JLabel;
import javax.swing.JOptionPane;
import javax.swing.SwingUtilities;
import javax.swing.UIManager;

import org.apache.poi.openxml4j.exceptions.NotOfficeXmlFileException;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

@SuppressWarnings("serial")
public class ExcelOnlyCount extends JFrame {

	private static ExcelOnlyCount instance = new ExcelOnlyCount();

	public static ExcelOnlyCount getInstance() {
		return instance;
	}

	// systemField
	int SystemFile1Sheet;
	int SystemFile2Sheet;
	int systemSheetNo1;
	int systemSheetNo2;

	// this is for accessing file 1 first row [ creating object ]
	FileInputStream file1;
	XSSFWorkbook workBook1;
	XSSFSheet sheet1;

	// this is for accessing file 2 first row [ creating object ]
	FileInputStream file2;
	XSSFWorkbook workBook2;
	XSSFSheet sheet2;

	// count ki liye
	private JLabel labelFILE1 = new JLabel("FILE 1 :");
	private JLabel labelFILE2 = new JLabel("FILE 2 :");
	private JLabel COUNT = new JLabel("COUNT1 :");
	private JLabel COUNT2 = new JLabel("COUNT2 :");
	private JLabel displayFileName1 = new JLabel();
	private JLabel displayFileName2 = new JLabel();
	private JComboBox<String> headerDropCount = new JComboBox<String>();
	private JComboBox<String> headerDropCount2 = new JComboBox<String>();
	private JButton buttonCount = new JButton("buttonCount");
	private JButton buttonCount2 = new JButton("buttonCount2");
	private JLabel selectSheet1 = new JLabel("SELECT SHEET 1 :");
	private JLabel selectSheet2 = new JLabel("SELECT SHEET 2 :");
	private JComboBox<String> selectSheet1Drop = new JComboBox<String>();
	private JComboBox<String> selectSheet2Drop2 = new JComboBox<String>();

	int selectedCounted;
	String selectedCountedName;
	int selectedCounted2;
	String selectedCountedName2;
	
	String filePath1Count;
	String filePath2Count;
	String targetFolderForCount;
	String fileName1ForCount;
	String fileName2ForCount;
//	String filePathForCount;
//	String filePath2ForCount;

	File filecreateFolder;
	Desktop desktop = Desktop.getDesktop();
	File file;

	private int countExcel(String filePath, String folderPath, String fileName, int selectedCounted, int selectedSheet,
			String selectedCountedName) {

		System.out.println("inside countExcel");

		System.out.println("filePath:" + filePath);
		System.out.println("folderPath:" + folderPath);
		System.out.println("fileName:" + fileName);
		System.out.println("selectedCounted:" + selectedCounted);
		System.out.println("selectedSheet:" + selectedSheet);
		System.out.println("selectedCountedName:" + selectedCountedName);

//		filecreateFolder = new File(folderPath + "\\Output");

		int count = 0;
		try {

			FileInputStream file1Count = new FileInputStream(filePath);
			XSSFWorkbook workBookCount = new XSSFWorkbook(file1Count);
			XSSFSheet sheetCount = workBookCount.getSheetAt(selectedSheet);

			int totalNumberOfRowsInExcel1Count = sheetCount.getLastRowNum();

			int columnIndex = selectedCounted;

			int total = 0;
			double totalP = 0;

			Set<String> set = new HashSet<>();

			for (int r = 1; r <= totalNumberOfRowsInExcel1Count; r++) {

				if (sheetCount.getRow(r) == null) {
					continue;
				}
				if (sheetCount.getRow(r).getCell(columnIndex) == null) {
					continue;
				}
				
				set.add(sheetCount.getRow(r).getCell(columnIndex).toString());

			}

			String[] setToStringArr = set.toArray(new String[set.size()]);
			int[] arr = new int[set.size()];

//			for (int i = 0; i < setToStringArr.length; i++) {
//				System.out.println(setToStringArr[i]);
//			}

			for (int i = 0; i < setToStringArr.length; i++) {

				for (int j = 1; j <= totalNumberOfRowsInExcel1Count; j++) {

					if (sheetCount.getRow(j) == null) {
						continue;
					}
					if (sheetCount.getRow(j).getCell(columnIndex) == null) {
						continue;
					}
					
					if (setToStringArr[i].equalsIgnoreCase(sheetCount.getRow(j).getCell(columnIndex).toString())) {
						arr[i]++;
					}
				}
			}

			for (int count1 : arr) {
//				System.out.println(count);
				total += count1;
//				totalP += count1;
			}

			// percentage
			double[] arrPer = new double[set.size()];
//			double percentage = 0;

			for (int count1 : arr) {
//				System.out.println(count1);
				totalP += count1;
			}

//			System.out.println("total:"+totalP);

			for (int i = 0; i < arrPer.length; i++) {
				arrPer[i] = (arr[i] / totalP) * 100;
			}

			// creating new working and adding new rows for excel1
			XSSFWorkbook workBookOutput1 = new XSSFWorkbook();
			XSSFSheet sheetCreate1 = workBookOutput1.createSheet();
			XSSFRow rowCreated = null;

			for (int r = 0; r <= setToStringArr.length; r++) {
				rowCreated = sheetCreate1.createRow(r);

				for (int c = 0; c < 4; c++) {
					rowCreated.createCell(c);
				}
			}

			for (int c = 0; c < 4; c++) {
				for (int i = 0; i <= setToStringArr.length; i++) {

					if (i < setToStringArr.length) {

						if (c == 0 && i == 0) {
							sheetCreate1.getRow(i).getCell(c).setCellValue(selectedCountedName);
						} else if (c == 1) {
							sheetCreate1.getRow(i).getCell(c).setCellValue(setToStringArr[i]);
						} else if (c == 2) {
							sheetCreate1.getRow(i).getCell(c).setCellValue(arr[i]);
						} else if (c == 3) {
							sheetCreate1.getRow(i).getCell(c).setCellValue(String.format("%.2f", arrPer[i]) + " %");
						}
					}

					if (i <= setToStringArr.length) {

						if (c == 1 && i == setToStringArr.length) {
							sheetCreate1.getRow(i).getCell(c).setCellValue("total:");
						} else if (c == 2 && i == setToStringArr.length) {
							sheetCreate1.getRow(i).getCell(c).setCellValue(total);
						}
					}
				}
			}

			String targetPathCount = folderPath + "\\Count_" + fileName;

			FileOutputStream outputStream11 = new FileOutputStream(targetPathCount);
			workBookOutput1.write(outputStream11);

			workBookOutput1.close();
			workBookCount.close();

			JOptionPane.showMessageDialog(ExcelOnlyCount.this, "Count Excel created", "Excel",
					JOptionPane.PLAIN_MESSAGE);
			System.out.println("Count1......Done");

		} catch (NullPointerException ne) {
//			count++;
			ne.printStackTrace();
		} catch (FileNotFoundException e1) {
			e1.printStackTrace();
//			count++;
		} catch (IOException ee) {
			ee.printStackTrace();
//			count++;
		}

		return count;

	}
	
//	private int countExcel(String filePath, String folderPath, String fileName, int selectedCounted, String selectedCountedName, int sheetNum) {
//
////		filecreateFolder = new File(folderPath +);
//
//		int count = 0;
//		try {
//
//			FileInputStream file1Count = new FileInputStream(filePath);
//			XSSFWorkbook workBookCount = new XSSFWorkbook(file1Count);
//			XSSFSheet sheetCount = workBookCount.getSheetAt(sheetNum);
//
//			int totalNumberOfRowsInExcel1Count = sheetCount.getLastRowNum();
//
//			int columnIndex = selectedCounted;
//
//			int total = 0;
//
//			Set<String> set = new HashSet<>();
//
//			for (int r = 1; r <= totalNumberOfRowsInExcel1Count; r++) {
//				
//				set.add(sheetCount.getRow(r).getCell(columnIndex).toString());
//			}
//
//			String[] setToStringArr = set.toArray(new String[set.size()]);
//			int[] arr = new int[set.size()];
//
////			for (int i = 0; i < setToStringArr.length; i++) {
////				System.out.println(setToStringArr[i]);
////			}
//
//			for (int i = 0; i < setToStringArr.length; i++) {
//
//				for (int j = 1; j <= totalNumberOfRowsInExcel1Count; j++) {
//
//					if (setToStringArr[i].equalsIgnoreCase(sheetCount.getRow(j).getCell(columnIndex).toString())) {
//						arr[i]++;
//					}
//				}
//			}
//
//			for (int count1 : arr) {
////				System.out.println(count);
//				total += count1;
//			}
//
//			// creating new working and adding new rows for excel1
//			XSSFWorkbook workBookOutput1 = new XSSFWorkbook();
//			XSSFSheet sheetCreate1 = workBookOutput1.createSheet();
//			XSSFRow rowCreated = null;
//
//			for (int r = 0; r <= setToStringArr.length; r++) {
//				rowCreated = sheetCreate1.createRow(r);
//
//				for (int c = 0; c < 3; c++) {
//					rowCreated.createCell(c);
//				}
//			}
//
//			for (int c = 0; c < 3; c++) {
//				for (int i = 0; i <= setToStringArr.length; i++) {
//
//					if (i < setToStringArr.length) {
//
//						if (c == 0 && i == 0) {
//							sheetCreate1.getRow(i).getCell(c).setCellValue(selectedCountedName);
//						} else if (c == 1) {
//							sheetCreate1.getRow(i).getCell(c).setCellValue(setToStringArr[i]);
//						} else if (c == 2) {
//							sheetCreate1.getRow(i).getCell(c).setCellValue(arr[i]);
//						}
//
//					}
//
//					if (i <= setToStringArr.length) {
//
//						if (c == 1 && i == setToStringArr.length) {
//							sheetCreate1.getRow(i).getCell(c).setCellValue("total:");
//						} else if (c == 2 && i == setToStringArr.length) {
//							sheetCreate1.getRow(i).getCell(c).setCellValue(total);
//						}
//
//					}
//				}
//			}
//
//			String targetPathCount = folderPath + "\\Count_" + fileName;
//
//			FileOutputStream outputStream11 = new FileOutputStream(targetPathCount);
//			workBookOutput1.write(outputStream11);
//
//			workBookOutput1.close();
//			workBookCount.close();
//
//			JOptionPane.showMessageDialog(ExcelOnlyCount.this, "Count Excel created", "Excel",
//					JOptionPane.PLAIN_MESSAGE);
//			System.out.println("Count1......Done");
//
//		} catch (NullPointerException ne) {
//			count++;
//			ne.printStackTrace();
//		} catch (FileNotFoundException e1) {
//			e1.printStackTrace();
//			count++;
//		} catch (IOException ee) {
//			ee.printStackTrace();
//			count++;
//		}
//
//		return count;
//	}

	private ExcelOnlyCount() {

//		swing started

		// setting title
		super("EXCEL COUNT");

		// setting layout
		setLayout(new GridBagLayout());

		GridBagConstraints constraints = new GridBagConstraints();
		constraints.anchor = GridBagConstraints.WEST;
		constraints.insets = new Insets(10, 10, 10, 10);

		// getting data from configuration
		String projectPath = System.getProperty("user.dir");

		try {

			File dir = new File(projectPath);
			String[] children = dir.list();
			
			targetFolderForCount = dir.toString() ;

			if (children == null) {
				System.out.println("does not exist or is not a directory");
			} else {
				boolean j = true;

				for (int i = 0; i < children.length; i++) {

					String fileName = children[i];

					if (fileName.length() > 5) {

						if (fileName.substring(fileName.length() - 5).equals(".xlsx") && j == true) {

							filePath1Count = projectPath + "\\" + fileName;
							j = false;
						}

						if (fileName.substring(fileName.length() - 5).equals(".xlsx") && j == false) {

							filePath2Count = projectPath + "\\" + fileName;

							if (filePath1Count.equals(filePath2Count)) {
								filePath2Count = null;
							}

						}
					}
				}
			}

//			System.out.println("filePath1Count:"+filePath1Count);
//			System.out.println("filePathCount:"+filePath2Count);
			
			
			File fileName1 = new File(filePath1Count);
			displayFileName1.setText("");
			if (fileName1.getName().length() < 12) {
				displayFileName1.setText(fileName1.getName());
			} else {
				displayFileName1.setText(fileName1.getName().substring(0, 12));
			}			
			
			File fileName2 = new File(filePath2Count);
			displayFileName2.setText("");
			if (fileName2.getName().length() < 12) {
				displayFileName2.setText(fileName2.getName());
			} else {
				displayFileName2.setText(fileName2.getName().substring(0, 12));
			}			
			
			try {

				selectSheet1Drop.removeAllItems();
				file1 = new FileInputStream(filePath1Count);
				workBook1 = new XSSFWorkbook(file1);

				int numberOfSheet1 = workBook1.getNumberOfSheets();

				for (int i = 0; i < numberOfSheet1; i++) {
					selectSheet1Drop.addItem(workBook1.getSheetName(i));
				}

				// main
				selectSheet1Drop.setSelectedIndex(SystemFile1Sheet);

				try {
					sheet1 = workBook1.getSheetAt(SystemFile1Sheet);
				} catch (IllegalArgumentException dd) {
				}

				if (sheet1.getRow(0) == null && sheet1.getRow(1) == null) {
					JOptionPane.showMessageDialog(ExcelOnlyCount.this, "Excel file 1 is Empty", "Excel",
							JOptionPane.ERROR_MESSAGE);
					filePath1Count = null;
				} else {
					int column = sheet1.getRow(0).getLastCellNum();
					XSSFRow row = sheet1.getRow(0);
					for (int c = 0; c < column; c++) {
						if (row.getCell(c) == null) {
							headerDropCount.addItem("");
						} else {
							headerDropCount.addItem("" + row.getCell(c));
						}
					}
				}
				
//				headerDropCount.setSelectedIndex(Systemkey1);

			} catch (NotOfficeXmlFileException e) {
			} catch (NullPointerException e) {
				JOptionPane.showMessageDialog(ExcelOnlyCount.this, "File 1 not found", "Excel",
						JOptionPane.PLAIN_MESSAGE);
			} catch (NumberFormatException e) {

				try {
					selectSheet1Drop.removeAllItems();
					file1 = new FileInputStream(filePath1Count);
					workBook1 = new XSSFWorkbook(file1);

					int numberOfSheet1 = workBook1.getNumberOfSheets();

					for (int i = 0; i < numberOfSheet1; i++) {
						selectSheet1Drop.addItem(workBook1.getSheetName(i));
					}

					try {
						sheet1 = workBook1.getSheetAt(SystemFile1Sheet);
					} catch (IllegalArgumentException dd) {
					}

					if (sheet1.getRow(0) == null) {
						JOptionPane.showMessageDialog(ExcelOnlyCount.this, "Excel file 1 is Empty", "Excel",
								JOptionPane.ERROR_MESSAGE);
						filePath1Count = null;
					} else {
						int column = sheet1.getRow(0).getLastCellNum();
						XSSFRow row = sheet1.getRow(0);
						for (int c = 0; c < column; c++) {
							if (row.getCell(c) == null) {
								headerDropCount.addItem("");
							} else {
								headerDropCount.addItem("" + row.getCell(c));
							}
						} // for

					}

				} catch (FileNotFoundException ee) {
					filePath1Count = "";
				}
			} catch (FileNotFoundException fee) {
			}
			
			
			try {

				selectSheet2Drop2.removeAllItems();
				file2 = new FileInputStream(filePath2Count);
				workBook2 = new XSSFWorkbook(file2);

				int numberOfSheet2 = workBook2.getNumberOfSheets();

				for (int i = 0; i < numberOfSheet2; i++) {
					selectSheet2Drop2.addItem(workBook2.getSheetName(i));
				}

				// main
				selectSheet2Drop2.setSelectedIndex(SystemFile2Sheet);

				try {
					sheet2 = workBook2.getSheetAt(SystemFile2Sheet);
				} catch (IllegalArgumentException dd) {
				}

				if (sheet2.getRow(0) == null) {
					JOptionPane.showMessageDialog(ExcelOnlyCount.this, "Excel file 2 is Empty", "Excel",
							JOptionPane.ERROR_MESSAGE);
					filePath2Count = null;
				} else {

					int column = sheet2.getRow(0).getLastCellNum();
					XSSFRow row = sheet2.getRow(0);

					for (int c = 0; c < column; c++) {
						if (row.getCell(c) == null) {
							headerDropCount2.addItem("");
						} else {
							headerDropCount2.addItem("" + row.getCell(c));
						}
					} // for
				}

//				headerDropCount2.setSelectedIndex(Systemkey2);

			} catch (NullPointerException e) {
				JOptionPane.showMessageDialog(ExcelOnlyCount.this, "File 2 not found", "Excel",
						JOptionPane.PLAIN_MESSAGE);
			} catch (NumberFormatException e) {
				try {
					selectSheet2Drop2.removeAllItems();
					file2 = new FileInputStream(filePath2Count);
					workBook2 = new XSSFWorkbook(file2);

					int numberOfSheet2 = workBook2.getNumberOfSheets();

					for (int i = 0; i < numberOfSheet2; i++) {
						selectSheet2Drop2.addItem(workBook2.getSheetName(i));
					}

					try {
						sheet2 = workBook2.getSheetAt(SystemFile2Sheet);
					} catch (IllegalArgumentException dd) {
					}

					if (sheet2.getRow(0) == null) {
						JOptionPane.showMessageDialog(ExcelOnlyCount.this, "Excel file 2 is Empty", "Excel",
								JOptionPane.ERROR_MESSAGE);
						filePath2Count = null;
					} else {

						int column = sheet2.getRow(0).getLastCellNum();
						XSSFRow row = sheet2.getRow(0);

						for (int c = 0; c < column; c++) {
							if (row.getCell(c) == null) {
								headerDropCount2.addItem("");
							} else {
								headerDropCount2.addItem("" + row.getCell(c));
							}
						} // for
					}

				} catch (FileNotFoundException ee) {
					filePath2Count = "";
				}
			} catch (FileNotFoundException e) {
			}

			try {
				File filePath1 = new File(filePath1Count);
				fileName1ForCount = filePath1.getName();
			} catch (Exception e) {
			}

			try {
				File filePath2 = new File(filePath2Count);
				fileName2ForCount = filePath2.getName();
			} catch (Exception e) {
			}

		} catch (FileNotFoundException fe) {
			JOptionPane.showMessageDialog(ExcelOnlyCount.this, "configuration file not found", "Excel",
					JOptionPane.ERROR_MESSAGE);
			fe.printStackTrace();
		} catch (IOException e1) {
			e1.printStackTrace();
		}
		
		constraints.gridy = 0;
		constraints.gridx = 0;
		add(labelFILE1, constraints);
		
		constraints.gridy = 1;
		constraints.gridx = 0;
		add(labelFILE2, constraints);
		
		constraints.gridy = 0;
		constraints.gridx = 1;
		add(displayFileName1, constraints);
		
		constraints.gridy = 1;
		constraints.gridx = 2;
		add(displayFileName2, constraints);
		
		constraints.gridy = 2;
		constraints.gridx = 0;
		add(selectSheet1, constraints);

		constraints.gridy = 2;
		constraints.gridx = 1;
		add(selectSheet1Drop, constraints);
		
		selectSheet1Drop.setPrototypeDisplayValue("XXXXXXXXXXXXXXXXXXX");
		selectSheet1Drop.setMaximumRowCount(10);

		selectSheet1Drop.addActionListener((e) -> {

			if (e.getSource() == selectSheet1Drop) {

				headerDropCount.removeAllItems();


				int selectedSheet1 = selectSheet1Drop.getSelectedIndex();
				systemSheetNo1 = selectedSheet1;
				SystemFile1Sheet = selectedSheet1;

				try {
					sheet1 = workBook1.getSheetAt(selectedSheet1);
				} catch (IllegalArgumentException dd) {
				}

//				sheetNo1 = selectedSheet1;
//				sheetName1 = sheet1.getSheetName();

				if (sheet1.getRow(0) == null) {
					JOptionPane.showMessageDialog(ExcelOnlyCount.this, "Excel file 1 is Empty", "Excel",
							JOptionPane.ERROR_MESSAGE);
//					path1 = null;
				} else {
					int column = sheet1.getRow(0).getLastCellNum();

					XSSFRow row = sheet1.getRow(0);
					for (int c = 0; c < column; c++) {
						if (row.getCell(c) == null) {
							headerDropCount.addItem("");

						} else {
							headerDropCount.addItem("" + row.getCell(c));

						}
					} // for

				}
			}
		});

		constraints.gridy = 3;
		constraints.gridx = 0;
		add(selectSheet2, constraints);

		constraints.gridy = 3;
		constraints.gridx = 1;
		add(selectSheet2Drop2, constraints);
		
		selectSheet2Drop2.setPrototypeDisplayValue("XXXXXXXXXXXXXXXXXXX");
		selectSheet2Drop2.setMaximumRowCount(10);

		selectSheet2Drop2.addActionListener((e) -> {
			if (e.getSource() == selectSheet2Drop2) {

				headerDropCount2.removeAllItems();
				
				int selectedSheet2 = selectSheet2Drop2.getSelectedIndex();
				systemSheetNo2 = selectedSheet2;
				SystemFile2Sheet = selectedSheet2;

				try {
					sheet2 = workBook2.getSheetAt(selectedSheet2);
				} catch (IllegalArgumentException dd) {

				}

//				sheetNo2 = selectedSheet2;
//				sheetName2 = sheet2.getSheetName();

				if (sheet2.getRow(0) == null) {
					JOptionPane.showMessageDialog(ExcelOnlyCount.this, "Excel file 2 is Empty", "Excel",
							JOptionPane.ERROR_MESSAGE);
//					path2 = null;
				} else {
					int column = sheet2.getRow(0).getLastCellNum();

					XSSFRow row = sheet2.getRow(0);

					for (int c = 0; c < column; c++) {
						if (row.getCell(c) == null) {
							headerDropCount2.addItem("");
						} else {
							headerDropCount2.addItem("" + row.getCell(c));
						}
					} // for
				}
			}
		});

		constraints.gridy = 4;
		constraints.gridx = 0;
		add(COUNT, constraints);

		constraints.gridy = 4;
		constraints.gridx = 1;
		add(headerDropCount, constraints);

		headerDropCount.setPrototypeDisplayValue("XXXXXXXXXXXXXXXXXXX");
		headerDropCount.setMaximumRowCount(10);

		headerDropCount.addActionListener((e) -> {
			if (e.getSource() == headerDropCount) {
				selectedCounted = headerDropCount.getSelectedIndex();
				selectedCountedName = (String) headerDropCount.getSelectedItem();
			}
		});

		constraints.gridy = 4;
		constraints.gridx = 2;
		add(buttonCount, constraints);

		buttonCount.addActionListener((e) -> {
			if (e.getSource() == buttonCount) {

//				System.out.println("filePathForCount:"+filePathForCount);

				int a = countExcel(filePath1Count, targetFolderForCount, fileName1ForCount, selectedCounted, systemSheetNo1,
						selectedCountedName);

				if (a == 0) {
				int ii = JOptionPane.showConfirmDialog(null,
						"We Have to close this window in order to open newly generated Excel, Because these are already open or are in use by javaw.exe Or if have to get more excels then click on No",
						"Exit?", JOptionPane.YES_NO_OPTION);
				if (ii == 1) {
					// do nothing
				}
				if (ii == 0) {
					try {
						File f = new File(targetFolderForCount);
						desktop.open(f);
					} catch (IOException eeee) {
						eeee.printStackTrace();
					}
					System.exit(0);
				}
				} else {
					JOptionPane.showMessageDialog(ExcelOnlyCount.this,
							"Excels creation NOT DONE/File is missing - Something is wrong!", "Excel !",
							JOptionPane.ERROR_MESSAGE);
				}
			}
		});

		constraints.gridy = 5;
		constraints.gridx = 0;
		add(COUNT2, constraints);

		constraints.gridy = 5;
		constraints.gridx = 1;
		add(headerDropCount2, constraints);

		headerDropCount2.setPrototypeDisplayValue("XXXXXXXXXXXXXXXXXXX");
		headerDropCount2.setMaximumRowCount(10);

		headerDropCount2.addActionListener((e) -> {
			if (e.getSource() == headerDropCount2) {
				selectedCounted2 = headerDropCount2.getSelectedIndex();
				selectedCountedName2 = (String) headerDropCount2.getSelectedItem();
			}
		});

		constraints.gridy = 5;
		constraints.gridx = 2;
		add(buttonCount2, constraints);

		buttonCount2.addActionListener((e) -> {

//			System.out.println("filePath2ForCount:" + filePath2ForCount);

			if (e.getSource() == buttonCount2) {

			int a = countExcel(filePath2Count, targetFolderForCount, fileName2ForCount, selectedCounted2, systemSheetNo2,
						selectedCountedName2);

				if (a == 0) {
					int ii = JOptionPane.showConfirmDialog(null,
							"We Have to close this window in order to open newly generated Excel, Because these are already open or are in use by javaw.exe Or if have to get more excels then click on No",
							"Exit?", JOptionPane.YES_NO_OPTION);
					if (ii == 1) {
						// do nothing
					}
					if (ii == 0) {
						try {
							File f = new File(targetFolderForCount);
							desktop.open(f);
						} catch (IOException eeee) {
							eeee.printStackTrace();
						}
						System.exit(0);
					}
				} else {
					JOptionPane.showMessageDialog(ExcelOnlyCount.this,
							"Excels creation NOT DONE/File is missing - Something is wrong!", "Excel !",
							JOptionPane.ERROR_MESSAGE);
				}

			}
		});

		pack();
		setResizable(false);
		setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
		setLocationRelativeTo(null);
	}

	public static void main(String[] args) {

		try {
			UIManager.setLookAndFeel(UIManager.getSystemLookAndFeelClassName());
		} catch (Exception e) {
		}

		SwingUtilities.invokeLater(new Runnable() {
			public void run() {
				getInstance().setVisible(true);
			}
		});
	}
}