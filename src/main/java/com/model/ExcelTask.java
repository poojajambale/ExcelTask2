package com.model;

import java.awt.Color;
import java.awt.Desktop;
import java.awt.GridBagConstraints;
import java.awt.GridBagLayout;
import java.awt.Insets;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.FileReader;
import java.io.IOException;
import java.util.Properties;

import javax.swing.Action;
import javax.swing.JButton;
import javax.swing.JComboBox;
import javax.swing.JFileChooser;
import javax.swing.JFrame;
import javax.swing.JLabel;
import javax.swing.JOptionPane;
import javax.swing.SwingUtilities;
import javax.swing.UIManager;
import javax.swing.filechooser.FileNameExtensionFilter;

import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.xssf.XLSBUnsupportedException;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

@SuppressWarnings("serial")
class Excel extends JFrame {

	private static Excel instance = new Excel();

	public static Excel getInstance() {
		return instance;
	}

	// outputName-1
	String fileName1;
	String fileName2;
	String sheetName1;
	String sheetName2;
	String keyName1;
	String keyName2;
	String folderName;

	// 53 line to 816 line all are for same data or duplicate data methods.
	public int duplicateExcel(String path1, String path2, int sheetNo1, int sheetNo2, int keyFile1, int keyFile2,
			String fileName1, String fileName2, String sheetName1, String sheetName2, String keyName1, String keyName2,
			String folderPath) {

		int countDup = 0;

		try {

			String firstExcelPath = path1;
			FileInputStream file1 = new FileInputStream(firstExcelPath);
			XSSFWorkbook workBook1 = new XSSFWorkbook(file1);
			XSSFSheet sheet1 = workBook1.getSheetAt(sheetNo1);

			String secondExcelPath = path2;
			FileInputStream file2 = new FileInputStream(secondExcelPath);
			XSSFWorkbook workBook2 = new XSSFWorkbook(file2);
			XSSFSheet sheet2 = workBook2.getSheetAt(sheetNo2);

			// workBook1
			int totalNumberOfRowsInExcel1 = sheet1.getLastRowNum();
			int totalNumberOfColumnInExcel1 = sheet1.getRow(0).getLastCellNum();

			// workBook2
			int totalNumberOfRowsInExcel2 = sheet2.getLastRowNum();
			int totalNumberOfColumnInExcel2 = sheet2.getRow(0).getLastCellNum();

			// going to Excel1 (key column) row => 1 to last
			for (int r = 1; r <= totalNumberOfRowsInExcel1; r++) {

				if (sheet1.getRow(r) == null) {
					continue;
				} else {

					int counterr = 0;

					if (sheet1.getRow(r).getCell(keyFile1) == null) {
						counterr = 0;
					} else {

						// going to Excel2 (key column) row => 1 to last
						for (int e = 1; e <= totalNumberOfRowsInExcel2; e++) {

							if (sheet2.getRow(e) == null) {
								continue;
							} else {

								if (sheet2.getRow(e).getCell(keyFile2) == null) {
									continue;
								}

								if (sheet1.getRow(r).getCell(keyFile1).toString()
										.equals(sheet2.getRow(e).getCell(keyFile2).toString())) {
//									System.out.println("SameCells1:" + sheet1.getRow(r).getCell(keyFile1) + "/"+ sheet2.getRow(e).getCell(keyFile2));
									counterr++;
									break;
								}
							}
						}
					} // else

					if (counterr == 0) {
						XSSFRow row = sheet1.getRow(r);
						sheet1.removeRow(row);
					}

				} // else
			} // for

			String firstExcelPathCopy = path1;
			FileInputStream file1Copy = new FileInputStream(firstExcelPathCopy);
			XSSFWorkbook workBook1Copy = new XSSFWorkbook(file1Copy);
			XSSFSheet sheet1Copy = workBook1Copy.getSheetAt(sheetNo1);

			// going to Excel2 key -> row = 1 to last
			for (int rr = 1; rr <= totalNumberOfRowsInExcel2; rr++) {
				if (sheet2.getRow(rr) == null) {
					continue;
				} else {

					int counterrr = 0;

					if (sheet2.getRow(rr).getCell(keyFile2) == null) {
						counterrr = 0;
					} else {

						// going to Excel1 key -> row = 1 to last
						for (int e = 1; e <= totalNumberOfRowsInExcel1; e++) {
							if (sheet1Copy.getRow(e) == null) {
								continue;
							} else {

								if (sheet1Copy.getRow(e).getCell(keyFile1) == null) {
									continue;
								}

								if (sheet2.getRow(rr).getCell(keyFile2).toString()
										.equals(sheet1Copy.getRow(e).getCell(keyFile1).toString())) {
//									System.out.println("SameCells2:" + sheet2.getRow(rr).getCell(keyFile2) + "/"+ sheet1Copy.getRow(e).getCell(keyFile1));
									counterrr++;
									break;
								}
							}
						} // for
					} // else

					if (counterrr == 0) {
//						XSSFRow row = sheet2.getRow(rr);
						sheet2.removeRow(sheet2.getRow(rr));
					}

				} // else
			} // for

			// upto hear we have same data but with null row

			// counting null row in EXCEL 1
			int counter = 0;
			for (int r = 1; r <= totalNumberOfRowsInExcel1; r++) {
				if (sheet1.getRow(r) == null) {
					counter++;
				}
			}

			// creating new working and adding new rows for excel1
			XSSFWorkbook workBookOutput1 = new XSSFWorkbook();
			XSSFSheet sheetCreate1 = workBookOutput1.createSheet();
			XSSFRow rowCreated = null;

			if (counter != 0) {

				// creating new excel 1 removing NULL row
				int totalNumberOfRowsOfNewSheet = totalNumberOfRowsInExcel1 - counter;

				for (int r = 0; r <= totalNumberOfRowsOfNewSheet; r++) {
					rowCreated = sheetCreate1.createRow(r);

					for (int c = 0; c < totalNumberOfColumnInExcel1; c++) {
						rowCreated.createCell(c);
					}
				}

				for (int p = 0, u = 0; p <= totalNumberOfRowsInExcel1; p++) {

					if (sheet1.getRow(p) == null) {
						continue;
					} else {

						rowCreated = sheetCreate1.getRow(u);

						for (int d = 0; d < totalNumberOfColumnInExcel1; d++) {

							if (sheet1.getRow(p).getCell(d) == null) {
								continue;
							} else {

								if (sheet1.getRow(p).getCell(d).getCellType() == CellType.STRING) {
									rowCreated.getCell(d)
											.setCellValue(sheet1.getRow(p).getCell(d).getStringCellValue());
								} else if (sheet1.getRow(p).getCell(d).getCellType() == CellType.NUMERIC) {
									rowCreated.getCell(d)
											.setCellValue(sheet1.getRow(p).getCell(d).getNumericCellValue());
								} else if (sheet1.getRow(p).getCell(d).getCellType() == CellType.BOOLEAN) {
									rowCreated.getCell(d)
											.setCellValue(sheet1.getRow(p).getCell(d).getBooleanCellValue());
								}

							}
						}
					}
					u++;
				}

				if (sheetCreate1.getLastRowNum() > 0) {
					try {
						// removed null excel writing
						System.out.println("Duplicate Excel1 created");

						String target1Path = folderPath + "\\Duplicate_ComparedBy_" + keyName1 + "_" + sheetName1 + "_"
								+ fileName1;
						FileOutputStream outputStream11 = new FileOutputStream(target1Path);
						workBookOutput1.write(outputStream11);
						workBookOutput1.close();
					} catch (FileNotFoundException ee) {
						countDup++;
						JOptionPane.showMessageDialog(Excel.this, "files 1 does'nt have Same data No Excel created",
								"Excel", JOptionPane.ERROR_MESSAGE);
					}
				} else {
					countDup++;
					JOptionPane.showMessageDialog(Excel.this, "files 1 does'nt have Same data No Excel created",
							"Excel", JOptionPane.ERROR_MESSAGE);
				}

			} // if counter close

			else {

				if (sheet1.getLastRowNum() > 0) {
					try {
						System.out.println("Duplicate Excel1 created");
						String target1Path1 = folderPath + "\\Duplicate_ComparedBy_" + keyName1 + "_" + sheetName1 + "_"
								+ fileName1;
						FileOutputStream outputStream1 = new FileOutputStream(target1Path1);
						workBook1.write(outputStream1);
					} catch (FileNotFoundException ee) {
						countDup++;
						JOptionPane.showMessageDialog(Excel.this, "files 1 does'nt have Same data No Excel created",
								"Excel", JOptionPane.ERROR_MESSAGE);
					}
				} else {
					countDup++;
					JOptionPane.showMessageDialog(Excel.this, "files 1 does'nt have Same data No Excel created",
							"Excel", JOptionPane.ERROR_MESSAGE);
				}
			}

			// counting null row in EXCEL 2
			int counter2 = 0;
			for (int r = 1; r <= totalNumberOfRowsInExcel2; r++) {
				if (sheet2.getRow(r) == null) {
					counter2++;
				}
			}

			// creating new working and adding new rows for excel2
			XSSFWorkbook workBookOutput2 = new XSSFWorkbook();
			XSSFSheet sheetCreate2 = workBookOutput2.createSheet();
			XSSFRow rowCreated2 = null;

			if (counter2 != 0) {

				int totalNumberOfRowsOfNewSheet2 = totalNumberOfRowsInExcel2 - counter2;

				for (int r = 0; r <= totalNumberOfRowsOfNewSheet2; r++) {
					rowCreated2 = sheetCreate2.createRow(r);

					for (int c = 0; c < totalNumberOfColumnInExcel2; c++) {
						rowCreated2.createCell(c);
					}
				}

				for (int p = 0, v = 0; p <= totalNumberOfRowsInExcel2; p++) {

					if (sheet2.getRow(p) == null) {
						continue;
					} else {
						rowCreated2 = sheetCreate2.getRow(v);

						for (int d = 0; d < totalNumberOfColumnInExcel2; d++) {

							if (sheet2.getRow(p).getCell(d) == null) {
								continue;
							} else {

								if (sheet2.getRow(p).getCell(d).getCellType() == CellType.STRING) {
									rowCreated2.getCell(d)
											.setCellValue(sheet2.getRow(p).getCell(d).getStringCellValue());
								} else if (sheet2.getRow(p).getCell(d).getCellType() == CellType.NUMERIC) {
									rowCreated2.getCell(d)
											.setCellValue(sheet2.getRow(p).getCell(d).getNumericCellValue());
								} else if (sheet2.getRow(p).getCell(d).getCellType() == CellType.BOOLEAN) {
									rowCreated2.getCell(d)
											.setCellValue(sheet2.getRow(p).getCell(d).getBooleanCellValue());
								}

							}
						}
					}
					v++;
				}

				// null row removed successfully
				// here we will have sheetCreate2

				if (sheetCreate2.getLastRowNum() > 0) {
					try {
						System.out.println("Duplicate Excel2 created");
						String target1Path = folderPath + "\\Duplicate_ComparedBy_" + keyName2 + "_" + sheetName2 + "_"
								+ fileName2;
						FileOutputStream outputStream22 = new FileOutputStream(target1Path);
						workBookOutput2.write(outputStream22);
						workBookOutput2.close();
					} catch (FileNotFoundException ee) {
						countDup++;
						JOptionPane.showMessageDialog(Excel.this, "files 2 does'nt have Same data No Excel created",
								"Excel", JOptionPane.ERROR_MESSAGE);
					}
				} else {
					countDup++;
					JOptionPane.showMessageDialog(Excel.this, "files 2 does'nt have Same data No Excel created",
							"Excel", JOptionPane.ERROR_MESSAGE);
				}

			} // if counter2 close

			else {
				if (sheet2.getLastRowNum() > 0) {
					try {
						System.out.println("Duplicate Excel2 created");
						String target2Path2 = folderPath + "\\Duplicate_ComparedBy_" + keyName2 + "_" + sheetName2 + "_"
								+ fileName2;
						FileOutputStream outputStream2 = new FileOutputStream(target2Path2);
						workBook2.write(outputStream2);
					} catch (FileNotFoundException ee) {
						countDup++;
						JOptionPane.showMessageDialog(Excel.this, "files 2 does'nt have Same data No Excel created",
								"Excel", JOptionPane.ERROR_MESSAGE);
					}
				} else {
					countDup++;
					JOptionPane.showMessageDialog(Excel.this, "files 2 does'nt have Same data No Excel created",
							"Excel", JOptionPane.ERROR_MESSAGE);
				}
			}

			workBook1Copy.close();
			workBook1.close();
			workBook2.close();

			System.out.println("Duplicate......Done");

		} catch (Exception e) {
		}

//			System.out.println("countDup:"+countDup);
		return countDup;

	}

	// for configuration file 2
	public int duplicateExcel(String path1, String path2, int sheetNo1, int sheetNo2, int keyFile1, int keyFile2,
			String fileName1, String fileName2, String folderPath) {

//		System.out.println("fileName1:"+fileName1);
//		System.out.println("fileName2:"+fileName2);

		int countDup = 0;

		try {
			String firstExcelPath = path1;
			FileInputStream file1 = new FileInputStream(firstExcelPath);
			XSSFWorkbook workBook1 = new XSSFWorkbook(file1);
			XSSFSheet sheet1 = workBook1.getSheetAt(sheetNo1);

			String secondExcelPath = path2;
			FileInputStream file2 = new FileInputStream(secondExcelPath);
			XSSFWorkbook workBook2 = new XSSFWorkbook(file2);
			XSSFSheet sheet2 = workBook2.getSheetAt(sheetNo2);

			// workBook1
			int totalNumberOfRowsInExcel1 = sheet1.getLastRowNum();
			int totalNumberOfColumnInExcel1 = sheet1.getRow(0).getLastCellNum();
//			XSSFCell cellOfRowKey1;

			// workBook2
			int totalNumberOfRowsInExcel2 = sheet2.getLastRowNum();
			int totalNumberOfColumnInExcel2 = sheet2.getRow(0).getLastCellNum();
//			XSSFCell cellOfRowKey2;

//			System.out.println("totalNumberOfRowsInExcel1:"+totalNumberOfRowsInExcel1);
//			System.out.println("totalNumberOfRowsInExcel2:"+totalNumberOfRowsInExcel2);

			// going to Excel1 key -> row = 1 to last
			for (int r = 1; r <= totalNumberOfRowsInExcel1; r++) {
				if (sheet1.getRow(r) == null) {
					continue;
				} else {

					int counterr = 0;

					if (sheet1.getRow(r).getCell(keyFile1) == null) {
//						continue;
						counterr = 0;
					} else {
//					    cellOfRowKey1 = sheet1.getRow(r).getCell(keyFile1);
//					}

//					System.out.println("cellOfRowKey1toString:" + cellOfRowKey1.toString());

						// going to Excel2 key -> row = 1 to last
						for (int e = 1; e <= totalNumberOfRowsInExcel2; e++) {
							if (sheet2.getRow(e) == null) {
								continue;
							} else {
								if (sheet2.getRow(e).getCell(keyFile2) == null) {
									continue;
								}
//								else {
//									cellOfRowKey2 = sheet2.getRow(e).getCell(keyFile2);
//								}

//							System.out.println("cellOfRowKey2toString:" + cellOfRowKey2.toString());

//							System.out.println(cellOfRowKey1.toString().equals(cellOfRowKey2.toString()));

								if (sheet1.getRow(r).getCell(keyFile1).toString()
										.equals(sheet2.getRow(e).getCell(keyFile2).toString())) {
//								System.out.println("SameCells1:" + cellOfRowKey1 + "/" + cellOfRowKey2);
//								XSSFRow row = sheet1.getRow(r);
//								countDup++;
									counterr++;
//								sheet1.removeRow(row);
//								continue;
									break;
								}
							}
						}
					}
					if (counterr == 0) {
//						XSSFRow row = sheet1.getRow(r);
						sheet1.removeRow(sheet1.getRow(r));
					}
//					continue;	
				}
			} // for

//			System.out.println("check");

//			System.out.println("-------------------------------------------------------------------------");

			String firstExcelPathCopy = path1;
			FileInputStream file1Copy = new FileInputStream(firstExcelPathCopy);
			XSSFWorkbook workBook1Copy = new XSSFWorkbook(file1Copy);
			XSSFSheet sheet1Copy = workBook1Copy.getSheetAt(sheetNo1);
//			XSSFCell cellOfRowKey1Copy;

			// going to Excel2 key -> row = 1 to last
			for (int rr = 1; rr <= totalNumberOfRowsInExcel2; rr++) {
				if (sheet2.getRow(rr) == null) {
					continue;
				} else {

					int counterrr = 0;

					if (sheet2.getRow(rr).getCell(keyFile2) == null) {
//						continue;
						counterrr = 0;
					} else {

//					if (sheet2.getRow(rr).getCell(keyFile2) == null) {
//						continue;
//					} else {
//						cellOfRowKey2 = sheet2.getRow(rr).getCell(keyFile2);
//					}
//					int counterrr = 0;
						// going to Excel1 key -> row = 1 to last
						for (int e = 1; e <= totalNumberOfRowsInExcel1; e++) {
							if (sheet1Copy.getRow(e) == null) {
								continue;
							} else {
								if (sheet1Copy.getRow(e).getCell(keyFile1) == null) {
									continue;
								} else {
//								cellOfRowKey1Copy = sheet1Copy.getRow(e).getCell(keyFile1);
								}

								if (sheet2.getRow(rr).getCell(keyFile2).toString()
										.equals(sheet1Copy.getRow(e).getCell(keyFile1).toString())) {
//								System.out.println("SameCells2:" + cellOfRowKey2 + "/" + cellOfRowKey1Copy);
//								countDup2++;
//								countDup++;
//								rowDup2.getCell(e).
//								XSSFRow row = sheet2.getRow(rr);
//								sheet2.removeRow(row);
									counterrr++;
									break;
//								continue;
								}
							}
						} // for
					}
					if (counterrr == 0) {
//						XSSFRow row = sheet2.getRow(rr);
						sheet2.removeRow(sheet2.getRow(rr));
					}
				}
			} // for

//			//read excel
//			System.out.println("excel2");
//			int totalNumberOfRows2 = sheet2.getLastRowNum();
//			int totalNumberOfColumn2 = sheet2.getRow(0).getLastCellNum();
//
//			System.out.println("totalNumberOfRows2:" + totalNumberOfRows2);
//			System.out.println("totalNumberOfColumn2:" + totalNumberOfColumn2);
//
//			for (int r = 0; r <= totalNumberOfRows2; r++) {
//				for (int c = 0; c < totalNumberOfColumn2; c++) {
//					if(sheet2.getRow(r) == null) {
//						continue;
//					} 
//					if (sheet2.getRow(r).getCell(c) == null) {
//						continue;
//					}
//					if (sheet2.getRow(r).getCell(c).getCellType() == CellType.STRING) {
//						System.out.print(sheet2.getRow(r).getCell(c).getStringCellValue());
//					} else if (sheet2.getRow(r).getCell(c).getCellType() == CellType.NUMERIC) {
//						System.out.print(sheet2.getRow(r).getCell(c).getNumericCellValue());
//					} else if (sheet2.getRow(r).getCell(c).getCellType() == CellType.BOOLEAN) {
//						System.out.print(sheet2.getRow(r).getCell(c).getBooleanCellValue());
//					}
//				}
//				System.out.println();
//			}
//
//			//excel read done

			// upto hear we have same data but with null row

			// counting null row in EXCEL 1
			int counter = 0;
			for (int r = 1; r <= totalNumberOfRowsInExcel1; r++) {
				if (sheet1.getRow(r) == null) {
					counter++;
				}
			}

//									System.out.println("counter:"+counter);

			// creating new working and adding new rows for excel1
			XSSFWorkbook workBookOutput1 = new XSSFWorkbook();
			XSSFSheet sheetCreate1 = workBookOutput1.createSheet();
			XSSFRow rowCreated = null;

//									if (false) {

			if (counter != 0) {

//												System.out.println("totalNumberOfRows1:" + totalNumberOfRowsInExcel1);
//												System.out.println("counter:" + counter);

				int totalNumberOfRowsOfNewSheet = totalNumberOfRowsInExcel1 - counter;

//												System.out.println("totalNumberOfRowsOfNewSheet1:" + totalNumberOfRowsOfNewSheet);

				for (int r = 0; r <= totalNumberOfRowsOfNewSheet; r++) {
					rowCreated = sheetCreate1.createRow(r);

					for (int c = 0; c < totalNumberOfColumnInExcel1; c++) {
//														cellCreated = rowCreated.createCell(c);
						rowCreated.createCell(c);
					}
				}

				for (int p = 0, u = 0; p <= totalNumberOfRowsInExcel1; p++) {
					if (sheet1.getRow(p) == null) {
						continue;
					} else {
						rowCreated = sheetCreate1.getRow(u);

//						for (int d = 0; d < totalNumberOfColumnInExcel1; d++) {
//							
//							if (sheet1.getRow(p).getCell(d) == null) {
//								continue;
//							} else {
//								if (sheet1.getRow(p).getCell(d).getCellType() == CellType.STRING) {
//									rowCreated.getCell(d)
//											.setCellValue(sheet1.getRow(p).getCell(d).getStringCellValue());
//								} else if (sheet1.getRow(p).getCell(d).getCellType() == CellType.NUMERIC) {
//									rowCreated.getCell(d)
//											.setCellValue(sheet1.getRow(p).getCell(d).getNumericCellValue());
//								} else if (sheet1.getRow(p).getCell(d).getCellType() == CellType.BOOLEAN) {
//									rowCreated.getCell(d)
//											.setCellValue(sheet1.getRow(p).getCell(d).getBooleanCellValue());
//								}
//							}
//						}

						for (int d = 0; d < totalNumberOfColumnInExcel1; d++) {

							if (sheet1.getRow(p).getCell(d) == null) {
								continue;
							} else {
								if (sheet1.getRow(p).getCell(d).getCellType() == CellType.STRING) {
									rowCreated.getCell(d)
											.setCellValue(sheet1.getRow(p).getCell(d).getStringCellValue());
								} else if (sheet1.getRow(p).getCell(d).getCellType() == CellType.NUMERIC) {
									rowCreated.getCell(d)
											.setCellValue(sheet1.getRow(p).getCell(d).getNumericCellValue());
								} else if (sheet1.getRow(p).getCell(d).getCellType() == CellType.BOOLEAN) {
									rowCreated.getCell(d)
											.setCellValue(sheet1.getRow(p).getCell(d).getBooleanCellValue());
								}
							}
						}
					}
					u++;
				}

//				System.out.println("NumberOfRow1:"+sheetCreate1.getLastRowNum());

				if (sheetCreate1.getLastRowNum() > 0) {
					try {
						// removed null excel writing
						System.out.println("Duplicate Excel1 created");
						String target1Path = folderPath + "\\Duplicate1_ComparedBy_" + fileName1;
						FileOutputStream outputStream11 = new FileOutputStream(target1Path);
						workBookOutput1.write(outputStream11);
						workBookOutput1.close();
					} catch (FileNotFoundException ee) {
						countDup++;
						JOptionPane.showMessageDialog(Excel.this, "files 1 does'nt have Same data No Excel created",
								"Excel", JOptionPane.ERROR_MESSAGE);
					}
				} else {
					countDup++;
					JOptionPane.showMessageDialog(Excel.this, "files 1 does'nt have Same data No Excel created",
							"Excel", JOptionPane.ERROR_MESSAGE);
				}

			} // if counter close

			else {

				if (sheet1.getLastRowNum() > 0) {
					try {
						System.out.println("Duplicate Excel1 created");
						String target1Path1 = folderPath + "\\Duplicate1_ComparedBy_" + fileName1;
						FileOutputStream outputStream1 = new FileOutputStream(target1Path1);
						workBook1.write(outputStream1);
					} catch (FileNotFoundException ee) {
						countDup++;
						JOptionPane.showMessageDialog(Excel.this, "files 1 does'nt have Same data No Excel created",
								"Excel", JOptionPane.ERROR_MESSAGE);
					}
				} else {
					countDup++;
					JOptionPane.showMessageDialog(Excel.this, "files 1 does'nt have Same data No Excel created",
							"Excel", JOptionPane.ERROR_MESSAGE);
				}
			}

//			System.out.println("countDup:"+countDup);

			// counting null row in EXCEL 2
			int counter2 = 0;
			for (int r = 1; r <= totalNumberOfRowsInExcel2; r++) {
				if (sheet2.getRow(r) == null) {
					counter2++;
				}
			}

//			System.out.println("counter2:"+counter2);

			// creating new working and adding new rows for excel2
			XSSFWorkbook workBookOutput2 = new XSSFWorkbook();
			XSSFSheet sheetCreate2 = workBookOutput2.createSheet();
			XSSFRow rowCreated2 = null;

			if (counter2 != 0) {

//						System.out.println("totalNumberOfRows2:" + totalNumberOfRowsInExcel2);
//						System.out.println("counter2:" + counter2);

				int totalNumberOfRowsOfNewSheet2 = totalNumberOfRowsInExcel2 - counter2;

//						System.out.println("totalNumberOfRowsOfNewSheet2:" + totalNumberOfRowsOfNewSheet2);

//							rowCreated2 = sheetCreate2.createRow(r);
//							rowCreated2.createCell(c);
//							rowCreated2 = sheetCreate2.getRow(v);

//						XSSFCell cellCreated2 = null;

				for (int r = 0; r <= totalNumberOfRowsOfNewSheet2; r++) {
					rowCreated2 = sheetCreate2.createRow(r);
					for (int c = 0; c < totalNumberOfColumnInExcel2; c++) {
//							cellCreated2 = rowCreated2.createCell(c);
						rowCreated2.createCell(c);
					}
				}

				for (int p = 0, v = 0; p <= totalNumberOfRowsInExcel2; p++) {
					if (sheet2.getRow(p) == null) {
						continue;
					} else {
						rowCreated2 = sheetCreate2.getRow(v);

						for (int d = 0; d < totalNumberOfColumnInExcel2; d++) {
							if (sheet2.getRow(p).getCell(d) == null) {
								continue;
							} else {
//									rowCreated2.getCell(d).setCellValue(sheet2.getRow(p).getCell(d).toString());
								if (sheet2.getRow(p).getCell(d).getCellType() == CellType.STRING) {
									rowCreated2.getCell(d)
											.setCellValue(sheet2.getRow(p).getCell(d).getStringCellValue());
								} else if (sheet2.getRow(p).getCell(d).getCellType() == CellType.NUMERIC) {
									rowCreated2.getCell(d)
											.setCellValue(sheet2.getRow(p).getCell(d).getNumericCellValue());
								} else if (sheet2.getRow(p).getCell(d).getCellType() == CellType.BOOLEAN) {
									rowCreated2.getCell(d)
											.setCellValue(sheet2.getRow(p).getCell(d).getBooleanCellValue());
								}
							}
						}
					}
					v++;
				}

				// null row removed successfully
				// here we will have to two sheetCreate1 and sheetCreate2

				if (sheetCreate2.getLastRowNum() > 0) {
					try {
						System.out.println("Duplicate Excel2 created");
						String target1Path = folderPath + "\\Duplicate2_ComparedBy_" + fileName2;
						FileOutputStream outputStream22 = new FileOutputStream(target1Path);
						workBookOutput2.write(outputStream22);
						workBookOutput2.close();
					} catch (FileNotFoundException ee) {
						countDup++;
						JOptionPane.showMessageDialog(Excel.this, "files 2 does'nt have Same data No Excel created",
								"Excel", JOptionPane.ERROR_MESSAGE);
					}
				} else {
					countDup++;
					JOptionPane.showMessageDialog(Excel.this, "files 2 does'nt have Same data No Excel created",
							"Excel", JOptionPane.ERROR_MESSAGE);
				}

			} // if counter close
			else {
				if (sheet2.getLastRowNum() > 0) {
					try {
						System.out.println("Duplicate Excel2 created");
						String target2Path2 = folderPath + "\\Duplicate2_ComparedBy_" + fileName2;
						FileOutputStream outputStream2 = new FileOutputStream(target2Path2);
						workBook2.write(outputStream2);
					} catch (FileNotFoundException ee) {
						countDup++;
						JOptionPane.showMessageDialog(Excel.this, "files 2 does'nt have Same data No Excel created",
								"Excel", JOptionPane.ERROR_MESSAGE);
					}
				} else {
					countDup++;
					JOptionPane.showMessageDialog(Excel.this, "files 2 does'nt have Same data No Excel created",
							"Excel", JOptionPane.ERROR_MESSAGE);
				}
			}

			workBook1Copy.close();
			workBook1.close();
			workBook2.close();

			System.out.println("Duplicate......Done");

		} catch (FileNotFoundException e) {
			countDup++;
			countDup++;
			JOptionPane.showMessageDialog(Excel.this, "file not found", "Excel", JOptionPane.ERROR_MESSAGE);
		} catch (Exception e) {
			countDup++;
		}

//		System.out.println("countDup:"+countDup);
		return countDup;
	}

	public int fetchExcel(String path1, String path2, int sheetNo1, int sheetNo2, int keyFile1, int keyFile2,
			String fileName1, String fileName2, String sheetName1, String sheetName2, String keyName1, String keyName2,
			String folderPath) {

		int counterMain = 0;

		try {

			String firstExcelPath = path1;
			FileInputStream file1 = new FileInputStream(firstExcelPath);
			XSSFWorkbook workBook1 = new XSSFWorkbook(file1);
			XSSFSheet sheet1 = workBook1.getSheetAt(sheetNo1);

			String secondExcelPath = path2;
			FileInputStream file2 = new FileInputStream(secondExcelPath);
			XSSFWorkbook workBook2 = new XSSFWorkbook(file2);
			XSSFSheet sheet2 = workBook2.getSheetAt(sheetNo2);

			// workBook1
			int totalNumberOfRowsInExcel1 = sheet1.getLastRowNum();
			int totalNumberOfColumnInExcel1 = sheet1.getRow(0).getLastCellNum();

			// workBook2
			int totalNumberOfRowsInExcel2 = sheet2.getLastRowNum();
			int totalNumberOfColumnInExcel2 = sheet2.getRow(0).getLastCellNum();
//			XSSFCell cellOfRowKey2;
//			XSSFRow rowOfSameKey2;

			// going to Excel1 key -> row = 1 to last
			for (int r = 1; r <= totalNumberOfRowsInExcel1; r++) {
				if (sheet1.getRow(r) == null) {
					continue;
				} else {
//					if (sheet1.getRow(r).getCell(keyFile1) == null) {
//						continue;
//					} else {
//						cellOfRowKey1 = sheet1.getRow(r).getCell(keyFile1);
//					}

//					int counterr = 0;

					if (sheet1.getRow(r).getCell(keyFile1) == null) {
						sheet1.removeRow(sheet1.getRow(r));
//						counterMain++;
						continue;
					}
//					else {

//					System.out.println("cellOfRowKey1:" + cellOfRowKey1);

					// going to Excel2 key -> row = 1 to last
					for (int e = 1; e <= totalNumberOfRowsInExcel2; e++) {
						if (sheet2.getRow(e) == null) {
							continue;
						} else {
							if (sheet2.getRow(e).getCell(keyFile2) == null) {
								continue;
							}
//							else {
//								cellOfRowKey2 = sheet2.getRow(e).getCell(keyFile2);
//							}
//							System.out.println("cellOfRowKey2:" + cellOfRowKey2);

							if ((sheet1.getRow(r).getCell(keyFile1).toString())
									.equals(sheet2.getRow(e).getCell(keyFile2).toString())) {
//								System.out.println("SameCells1:" + cellOfRowKey1 + "/" + cellOfRowKey2);
//								countDup1++;
								XSSFRow rowOfSameKey1 = sheet1.getRow(r);
								sheet1.removeRow(rowOfSameKey1);
//								sheet1.removeRow(sheet1.getRow(r));
								break;
							}
						}
					}
				}
			} // for

//			System.out.println("-------------------------------------------------------------------------");

			String firstExcelPathCopy = path1;
//			String firstExcelPathCopy = "C:\\Users\\SATYASAH\\Downloads\\Capg Bench.xlsx";
			FileInputStream file1Copy = new FileInputStream(firstExcelPathCopy);
			XSSFWorkbook workBook1Copy = new XSSFWorkbook(file1Copy);
			XSSFSheet sheet1Copy = workBook1Copy.getSheetAt(sheetNo1);
//			XSSFCell cellOfRowKey1Copy;

			// going to Excel2 key -> row = 1 to last
			for (int rr = 1; rr <= totalNumberOfRowsInExcel2; rr++) {
				if (sheet2.getRow(rr) == null) {
					continue;
				} else {
					if (sheet2.getRow(rr).getCell(keyFile2) == null) {
						sheet2.removeRow(sheet2.getRow(rr));
						continue;
					}
//					else {
//						sheet2.getRow(rr).getCell(keyFile2);
//					}
					// going to Excel1 key -> row = 1 to last
					for (int e = 1; e <= totalNumberOfRowsInExcel1; e++) {
						if (sheet1Copy.getRow(e) == null) {
							continue;
						} else {
							if (sheet1Copy.getRow(e).getCell(keyFile1) == null) {
								continue;
							}
//							else {
//								cellOfRowKey1Copy = sheet1Copy.getRow(e).getCell(keyFile1);
//							}

							if (sheet2.getRow(rr).getCell(keyFile2).toString()
									.equals(sheet1Copy.getRow(e).getCell(keyFile1).toString())) {
//								System.out.println("SameCells2:" + cellOfRowKey2 + "/" + cellOfRowKey1Copy);
//								countDup2++;
//								rowDup2.getCell(e).
//								rowOfSameKey2 = sheet2.getRow(rr);
								sheet2.removeRow(sheet2.getRow(rr));
								break;
							}
						}
					} // for
				}
			} // for
			workBook1Copy.close();
			// Upto here we have to two excel with some null or empty row
			// sheet1 and sheet2 as output only NO new sheet created

//-----------------------------------------------------------------------------------------------------------------			

			// counting null row in EXCEL 1
			int counter = 0;
			for (int rq = 0; rq <= totalNumberOfRowsInExcel1; rq++) {
				if (sheet1.getRow(rq) == null) {
					counter++;
				}
			}

//			System.out.println("totalNumberOfRows1:" + totalNumberOfRowsInExcel1);
//			System.out.println("counter:" + counter);

			// creating new working and adding new rows for excel1
			XSSFWorkbook workBookOutput1 = new XSSFWorkbook();
			XSSFSheet sheetCreate1 = workBookOutput1.createSheet();
			XSSFRow rowCreated = null;

			if (counter != 0) {

				int totalNumberOfRowsOfNewSheet = totalNumberOfRowsInExcel1 - counter;

//			System.out.println("totalNumberOfRowsOfNewSheet1:" + totalNumberOfRowsOfNewSheet);

				for (int rr = 0; rr <= totalNumberOfRowsOfNewSheet; rr++) {
					rowCreated = sheetCreate1.createRow(rr);

					for (int c = 0; c < totalNumberOfColumnInExcel1; c++) {
//					 XSSFCell cellCreated = rowCreated.createCell(c);
						rowCreated.createCell(c);
//						sheetCreate1.createRow(rr).createCell(c);
					}
				}

				for (int p = 0, u = 0; p <= totalNumberOfRowsInExcel1; p++) {
					if (sheet1.getRow(p) == null) {
						continue;
					} else {
						rowCreated = sheetCreate1.getRow(u);

						for (int d = 0; d < totalNumberOfColumnInExcel1; d++) {
							if (sheet1.getRow(p).getCell(d) == null) {
								continue;
							} else {
								if (sheet1.getRow(p).getCell(d).getCellType() == CellType.STRING) {
									rowCreated.getCell(d)
											.setCellValue(sheet1.getRow(p).getCell(d).getStringCellValue());
								} else if (sheet1.getRow(p).getCell(d).getCellType() == CellType.NUMERIC) {
									rowCreated.getCell(d)
											.setCellValue(sheet1.getRow(p).getCell(d).getNumericCellValue());
								} else if (sheet1.getRow(p).getCell(d).getCellType() == CellType.BOOLEAN) {
									rowCreated.getCell(d)
											.setCellValue(sheet1.getRow(p).getCell(d).getBooleanCellValue());
								}
							}
						}
						u++;
					}
				}

//				System.out.println("NumberOfRow1:"+sheetCreate1.getLastRowNum());
//				if (sheetCreate1.getLastRowNum() > 0) {
				try {
					System.out.println("Unique Excel1 created");
					String target1Path = folderPath + "\\ChildOutput_ComparedBy_" + keyName1 + "_" + sheetName1 + "_"
							+ fileName1;
					FileOutputStream outputStream11 = new FileOutputStream(target1Path);
					workBookOutput1.write(outputStream11);
					workBookOutput1.close();
				} catch (FileNotFoundException ee) {
					counterMain++;
					JOptionPane.showMessageDialog(Excel.this, "files 1 does'nt have unique data No Excel created",
							"Excel", JOptionPane.ERROR_MESSAGE);
				}
//				} else {
//					counterMain++;
//					JOptionPane.showMessageDialog(Excel.this, "files 1 does'nt have unique data No Excel created",
//							"Excel", JOptionPane.ERROR_MESSAGE);
//				}

			} // if counter close

			else {

//				System.out.println("sheet1.getLastRowNum():"+sheet1.getLastRowNum());

//				if (sheet1.getLastRowNum() > 0) {
				try {
					System.out.println("Unique Excel1 created");
					String target1Path1 = folderPath + "\\ChildOutput_ComparedBy_" + keyName1 + "_" + sheetName1 + "_"
							+ fileName1;
					FileOutputStream outputStream1 = new FileOutputStream(target1Path1);
					workBook1.write(outputStream1);
					workBook1.close();
				} catch (FileNotFoundException ee) {
					counterMain++;
					JOptionPane.showMessageDialog(Excel.this, "files 1 does'nt have unique data No Excel created",
							"Excel", JOptionPane.ERROR_MESSAGE);
				}
			}

			// counting null row in EXCEL 2
			int counter2 = 0;
			for (int r1 = 1; r1 <= totalNumberOfRowsInExcel2; r1++) {
				if (sheet2.getRow(r1) == null) {
					counter2++;
				}
			}

			// creating new working and adding new rows for excel2
			XSSFWorkbook workBookOutput2 = new XSSFWorkbook();
			XSSFSheet sheetCreate2 = workBookOutput2.createSheet();
			XSSFRow rowCreated2 = null;

			if (counter2 != 0) {

				int totalNumberOfRowsOfNewSheet2 = totalNumberOfRowsInExcel2 - counter2;

//			System.out.println("totalNumberOfRowsOfNewSheet2:" + totalNumberOfRowsOfNewSheet2);

//				rowCreated2 = sheetCreate2.createRow(r);
//				rowCreated2.createCell(c);
//				rowCreated2 = sheetCreate2.getRow(v);

//			XSSFCell cellCreated2 = null;

				for (int r2 = 0; r2 <= totalNumberOfRowsOfNewSheet2; r2++) {
					rowCreated2 = sheetCreate2.createRow(r2);
					for (int c = 0; c < totalNumberOfColumnInExcel2; c++) {
//				cellCreated2 = rowCreated2.createCell(c);
						rowCreated2.createCell(c);
					}
				}

				for (int p = 0, v = 0; p <= totalNumberOfRowsInExcel2; p++) {
					if (sheet2.getRow(p) == null) {
						continue;
					} else {
						rowCreated2 = sheetCreate2.getRow(v);
						for (int d = 0; d < totalNumberOfColumnInExcel2; d++) {
							if (sheet2.getRow(p).getCell(d) == null) {
								continue;
							} else {
//								rowCreated2.getCell(d).setCellValue(sheet2.getRow(p).getCell(d).toString());
								if (sheet2.getRow(p).getCell(d).getCellType() == CellType.STRING) {
									rowCreated2.getCell(d)
											.setCellValue(sheet2.getRow(p).getCell(d).getStringCellValue());
								} else if (sheet2.getRow(p).getCell(d).getCellType() == CellType.NUMERIC) {
									rowCreated2.getCell(d)
											.setCellValue(sheet2.getRow(p).getCell(d).getNumericCellValue());
								} else if (sheet2.getRow(p).getCell(d).getCellType() == CellType.BOOLEAN) {
									rowCreated2.getCell(d)
											.setCellValue(sheet2.getRow(p).getCell(d).getBooleanCellValue());
								}
							}
						}
					}
					v++;
				}

				// null row removed successfully
				// here we will have to two sheetCreate1 and sheetCreate2

//				System.out.println("NumberOfRow2:"+sheetCreate2.getLastRowNum());
//				if (sheetCreate2.getLastRowNum() > 0) {
				try {
					System.out.println("Unique Excel2 created");
					String target1Path = folderPath + "\\ChildOutput_ComparedBy_" + keyName2 + "_" + sheetName2 + "_"
							+ fileName2;
					FileOutputStream outputStream22 = new FileOutputStream(target1Path);
					workBookOutput2.write(outputStream22);
					workBookOutput2.close();
				} catch (FileNotFoundException ee) {
					counterMain++;
					JOptionPane.showMessageDialog(Excel.this, "files 2 does'nt have unique data No Excel created",
							"Excel", JOptionPane.ERROR_MESSAGE);
				}
//				} else {
//					counterMain++;
//					JOptionPane.showMessageDialog(Excel.this, "files 2 does'nt have unique data No Excel created",
//							"Excel", JOptionPane.ERROR_MESSAGE);
//				}
			} // if counter close
			else {
//				if (sheet2.getLastRowNum() > 0) {
				try {
					System.out.println("Unique Excel2 created");
					String target2Path2 = folderPath + "\\ChildOutput_ComparedBy_" + keyName2 + "_" + sheetName2 + "_"
							+ fileName2;
					FileOutputStream outputStream2 = new FileOutputStream(target2Path2);
					workBook2.write(outputStream2);
					workBook2.close();
				} catch (FileNotFoundException ee) {
					counterMain++;
					JOptionPane.showMessageDialog(Excel.this, "files 2 does'nt have unique data No Excel created",
							"Excel", JOptionPane.ERROR_MESSAGE);
				}
//				} else {
//					counterMain++;
//					JOptionPane.showMessageDialog(Excel.this, "files 2 does'nt have unique data No Excel created",
//							"Excel", JOptionPane.ERROR_MESSAGE);
//				}
			}

			
			workBook1Copy.close();

//-----------------------------------------------------------------------------------------------------------

//			// upto unique data withOut Null row Completed
//			if (countDup1 != 0) {
//
//			} else {
//
//			}

			System.out.println("Unique......Done");

		} catch (Exception e) {
			e.printStackTrace();
		}

		return counterMain;
	} // end of fetch method

	// for configuration file 2
	public int fetchExcel(String path1, String path2, int sheetNo1, int sheetNo2, int keyFile1, int keyFile2,
			String fileName1, String fileName2, String folderPath) {

		int counterMain = 0;
		try {

			String firstExcelPath = path1;
			FileInputStream file1 = new FileInputStream(firstExcelPath);
			XSSFWorkbook workBook1 = new XSSFWorkbook(file1);
			XSSFSheet sheet1 = workBook1.getSheetAt(sheetNo1);

			String secondExcelPath = path2;
			FileInputStream file2 = new FileInputStream(secondExcelPath);
			XSSFWorkbook workBook2 = new XSSFWorkbook(file2);
			XSSFSheet sheet2 = workBook2.getSheetAt(sheetNo2);

			// workBook1
			int totalNumberOfRowsInExcel1 = sheet1.getLastRowNum();
			int totalNumberOfColumnInExcel1 = sheet1.getRow(0).getLastCellNum();
//			XSSFCell cellOfRowKey1;
//			XSSFRow rowOfSameKey1;

			// workBook2
			int totalNumberOfRowsInExcel2 = sheet2.getLastRowNum();
			int totalNumberOfColumnInExcel2 = sheet2.getRow(0).getLastCellNum();
//			XSSFCell cellOfRowKey2;
//			XSSFRow rowOfSameKey2;

			// going to Excel1 key -> row = 1 to last
			for (int r = 1; r <= totalNumberOfRowsInExcel1; r++) {
				if (sheet1.getRow(r) == null) {
					continue;
				} else {
					if (sheet1.getRow(r).getCell(keyFile1) == null) {
						sheet1.removeRow(sheet1.getRow(r));
						continue;
					}
//					else {
//						cellOfRowKey1 = sheet1.getRow(r).getCell(keyFile1);
//					}

//					System.out.println("cellOfRowKey1:" + cellOfRowKey1);

					// going to Excel2 key -> row = 1 to last
					for (int e = 1; e <= totalNumberOfRowsInExcel2; e++) {
						if (sheet2.getRow(e) == null) {
							continue;
						} else {
							if (sheet2.getRow(e).getCell(keyFile2) == null) {
								continue;
							}
//							else {
//								cellOfRowKey2 = sheet2.getRow(e).getCell(keyFile2);
//							}
//							System.out.println("cellOfRowKey2:" + cellOfRowKey2);

							if ((sheet1.getRow(r).getCell(keyFile1).toString())
									.equals(sheet2.getRow(e).getCell(keyFile2).toString())) {
//								System.out.println("SameCells1:" + cellOfRowKey1 + "/" + cellOfRowKey2);
//								countDup1++;
								XSSFRow rowOfSameKey1 = sheet1.getRow(r);
								sheet1.removeRow(rowOfSameKey1);
//								sheet1.removeRow(sheet1.getRow(r));
								break;
							}
						}
					}
				}
			} // for

//			System.out.println("-------------------------------------------------------------------------");

			String firstExcelPathCopy = path1;
//			String firstExcelPathCopy = "C:\\Users\\SATYASAH\\Downloads\\Capg Bench.xlsx";
			FileInputStream file1Copy = new FileInputStream(firstExcelPathCopy);
			XSSFWorkbook workBook1Copy = new XSSFWorkbook(file1Copy);
			XSSFSheet sheet1Copy = workBook1Copy.getSheetAt(sheetNo1);
//			XSSFCell cellOfRowKey1Copy;

			// going to Excel2 key -> row = 1 to last
			for (int rr = 1; rr <= totalNumberOfRowsInExcel2; rr++) {
				if (sheet2.getRow(rr) == null) {
					continue;
				} else {
					if (sheet2.getRow(rr).getCell(keyFile2) == null) {
						sheet2.removeRow(sheet2.getRow(rr));
						continue;
					}
//					else {
//						cellOfRowKey2 = sheet2.getRow(rr).getCell(keyFile2);
//					}
					// going to Excel1 key -> row = 1 to last
					for (int e = 1; e <= totalNumberOfRowsInExcel1; e++) {
						if (sheet1Copy.getRow(e) == null) {
							continue;
						} else {
							if (sheet1Copy.getRow(e).getCell(keyFile1) == null) {
								continue;
							}
//							else {
//								cellOfRowKey1Copy = sheet1Copy.getRow(e).getCell(keyFile1);
//							}

							if (sheet2.getRow(rr).getCell(keyFile2).toString()
									.equals(sheet1Copy.getRow(e).getCell(keyFile1).toString())) {
//								System.out.println("SameCells2:" + cellOfRowKey2 + "/" + cellOfRowKey1Copy);
//								countDup2++;
//								rowDup2.getCell(e).
//								rowOfSameKey2 = sheet2.getRow(rr);
								sheet2.removeRow(sheet2.getRow(rr));
								break;
							}
						}
					} // for
				}
			} // for

			// Upto here we have to two excel with some null or empty row
			// sheet1 and sheet2 as output only NO new sheet created

//-----------------------------------------------------------------------------------------------------------------			

			// counting null row in EXCEL 1
			int counter = 0;
			for (int r = 0; r <= totalNumberOfRowsInExcel1; r++) {
				if (sheet1.getRow(r) == null) {
					counter++;
				}
			}

			// creating new working and adding new rows for excel1
			XSSFWorkbook workBookOutput1 = new XSSFWorkbook();
			XSSFSheet sheetCreate1 = workBookOutput1.createSheet();
			XSSFRow rowCreated = null;

			if (counter != 0) {

//			System.out.println("totalNumberOfRows1:" + totalNumberOfRowsInExcel1);
//			System.out.println("counter:" + counter);

				int totalNumberOfRowsOfNewSheet = totalNumberOfRowsInExcel1 - counter;

//			System.out.println("totalNumberOfRowsOfNewSheet1:" + totalNumberOfRowsOfNewSheet);

				for (int r = 0; r <= totalNumberOfRowsOfNewSheet; r++) {
					rowCreated = sheetCreate1.createRow(r);

					for (int c = 0; c < totalNumberOfColumnInExcel1; c++) {
//					cellCreated = rowCreated.createCell(c);
						rowCreated.createCell(c);
					}
				}

				for (int p = 0, u = 0; p <= totalNumberOfRowsInExcel1; p++) {
					if (sheet1.getRow(p) == null) {
						continue;
					} else {
						rowCreated = sheetCreate1.getRow(u);

						for (int d = 0; d < totalNumberOfColumnInExcel1; d++) {
							if (sheet1.getRow(p).getCell(d) == null) {
								continue;
							} else {
								if (sheet1.getRow(p).getCell(d).getCellType() == CellType.STRING) {
									rowCreated.getCell(d)
											.setCellValue(sheet1.getRow(p).getCell(d).getStringCellValue());
								} else if (sheet1.getRow(p).getCell(d).getCellType() == CellType.NUMERIC) {
									rowCreated.getCell(d)
											.setCellValue(sheet1.getRow(p).getCell(d).getNumericCellValue());
								} else if (sheet1.getRow(p).getCell(d).getCellType() == CellType.BOOLEAN) {
									rowCreated.getCell(d)
											.setCellValue(sheet1.getRow(p).getCell(d).getBooleanCellValue());
								}
							}
						}
						u++;
					}
				}

				try {
					System.out.println("Unique Excel1 created");
					String target1Path = folderPath + "\\ChildOutput_ComparedBy_" + fileName1;
					FileOutputStream outputStream11 = new FileOutputStream(target1Path);
					workBookOutput1.write(outputStream11);
					workBookOutput1.close();
				} catch (FileNotFoundException ee) {
					counterMain++;
					JOptionPane.showMessageDialog(Excel.this, "files 1 does'nt have unique data No Excel created",
							"Excel", JOptionPane.ERROR_MESSAGE);
				}

			} // if counter close

			else {

				try {
					System.out.println("Unique Excel1 created");
					String target1Path1 = folderPath + "\\ChildOutput_ComparedBy_" + fileName1;
					FileOutputStream outputStream1 = new FileOutputStream(target1Path1);
					workBook1.write(outputStream1);
					workBook1.close();
				} catch (FileNotFoundException ee) {
					counterMain++;
					JOptionPane.showMessageDialog(Excel.this, "files 1 does'nt have unique data No Excel created",
							"Excel", JOptionPane.ERROR_MESSAGE);
				}
			}

			// counting null row in EXCEL 2
			int counter2 = 0;
			for (int r = 1; r <= totalNumberOfRowsInExcel2; r++) {
				if (sheet2.getRow(r) == null) {
					counter2++;
				}
			}

			// creating new working and adding new rows for excel2
			XSSFWorkbook workBookOutput2 = new XSSFWorkbook();
			XSSFSheet sheetCreate2 = workBookOutput2.createSheet();
			XSSFRow rowCreated2 = null;

			if (counter2 != 0) {

//			System.out.println("totalNumberOfRows2:" + totalNumberOfRowsInExcel2);
//			System.out.println("counter2:" + counter2);

				int totalNumberOfRowsOfNewSheet2 = totalNumberOfRowsInExcel2 - counter2;

//			System.out.println("totalNumberOfRowsOfNewSheet2:" + totalNumberOfRowsOfNewSheet2);

//				rowCreated2 = sheetCreate2.createRow(r);
//				rowCreated2.createCell(c);
//				rowCreated2 = sheetCreate2.getRow(v);

//			XSSFCell cellCreated2 = null;

				for (int r = 0; r <= totalNumberOfRowsOfNewSheet2; r++) {
					rowCreated2 = sheetCreate2.createRow(r);
					for (int c = 0; c < totalNumberOfColumnInExcel2; c++) {
//				cellCreated2 = rowCreated2.createCell(c);
						rowCreated2.createCell(c);
					}
				}

				for (int p = 0, v = 0; p <= totalNumberOfRowsInExcel2; p++) {
					if (sheet2.getRow(p) == null) {
						continue;
					} else {
						rowCreated2 = sheetCreate2.getRow(v);
						for (int d = 0; d < totalNumberOfColumnInExcel2; d++) {
							if (sheet2.getRow(p).getCell(d) == null) {
								continue;
							} else {
//									rowCreated2.getCell(d).setCellValue(sheet2.getRow(p).getCell(d).toString());
								if (sheet2.getRow(p).getCell(d).getCellType() == CellType.STRING) {
									rowCreated2.getCell(d)
											.setCellValue(sheet2.getRow(p).getCell(d).getStringCellValue());
								} else if (sheet2.getRow(p).getCell(d).getCellType() == CellType.NUMERIC) {
									rowCreated2.getCell(d)
											.setCellValue(sheet2.getRow(p).getCell(d).getNumericCellValue());
								} else if (sheet2.getRow(p).getCell(d).getCellType() == CellType.BOOLEAN) {
									rowCreated2.getCell(d)
											.setCellValue(sheet2.getRow(p).getCell(d).getBooleanCellValue());
								}
							}
						}
					}
					v++;
				}

				// null row removed successfully
				// here we will have to two sheetCreate1 and sheetCreate2

//				System.out.println("NumberOfRow2:"+sheetCreate2.getLastRowNum());
//				if (sheetCreate2.getLastRowNum() > 0) {

				try {
					System.out.println("Unique Excel2 created");
					String target1Path = folderPath + "\\ChildOutput2_ComparedBy_" + fileName2;
					FileOutputStream outputStream22 = new FileOutputStream(target1Path);
					workBookOutput2.write(outputStream22);
					workBookOutput2.close();
				} catch (FileNotFoundException ee) {
					counterMain++;
					JOptionPane.showMessageDialog(Excel.this, "files 2 does'nt have unique data No Excel created",
							"Excel", JOptionPane.ERROR_MESSAGE);
				}

			} // if counter close
			else {

//				if (sheet2.getLastRowNum() > 0) {
				try {
					System.out.println("Unique Excel2 created");
					String target2Path2 = folderPath + "\\ChildOutput2_ComparedBy_" + fileName2;
					FileOutputStream outputStream2 = new FileOutputStream(target2Path2);
					workBook2.write(outputStream2);
					workBook2.close();
				} catch (FileNotFoundException ee) {
					counterMain++;
					JOptionPane.showMessageDialog(Excel.this, "files 2 does'nt have unique data No Excel created",
							"Excel", JOptionPane.ERROR_MESSAGE);
				}
			}

			workBook1Copy.close();
			

//			// upto unique data withOut Null row Completed
//			if (countDup1 != 0) {
//
//			} else {
//
//			}

			System.out.println("Unique......Done");

		} catch (FileNotFoundException fe) {
			counterMain++;
			counterMain++;
			JOptionPane.showMessageDialog(Excel.this, "file not found", "Excel", JOptionPane.ERROR_MESSAGE);
		} catch (Exception e) {
			counterMain++;
			e.printStackTrace();
		}

		return counterMain;
	}

//	swing started
//	by Manu, satyam and sangmesh
//---------------------------------------------------------------------------
	
	// class field
	// these are the child components for jFrame
	private JLabel labelFILE1 = new JLabel("FILE 1 :");
	private JLabel labelFILE2 = new JLabel("FILE 2 :");
	private JLabel labelKEYFILE1 = new JLabel("KEY 1 :");
	private JLabel labelKEYFILE2 = new JLabel("KEY 2 :");

	private JLabel outputFolder = new JLabel("OUTPUT :");
	private JLabel selectSheet1 = new JLabel("SELECT SHEET 1 :");
	private JLabel selectSheet2 = new JLabel("SELECT SHEET 2 :");
	private JComboBox<String> selectSheet1Drop = new JComboBox<String>();
	private JComboBox<String> selectSheet2Drop2 = new JComboBox<String>();
	private JLabel displayFileName1 = new JLabel();
	private JLabel displayFileName2 = new JLabel();
	private JLabel displayOutputFolder = new JLabel();
	private JComboBox<String> headerDrop = new JComboBox<String>();
	private JComboBox<String> headerDrop2 = new JComboBox<String>();
	private JButton buttonFile1 = new JButton("openFile1");
	private JButton buttonFile2 = new JButton("openFile2");
	private JButton buttonOutput = new JButton("openFolder");
	private JButton buttonENTER = new JButton("ENTER FOR UNIQUE");
	private JButton buttonDUPLICATE = new JButton("ENTER FOR DUPLICATE");
	private JButton buttonClear = new JButton("CLEAR");
	private File file;
	Desktop desktop = Desktop.getDesktop();

	// this is for accessing file 1 first row [ creating object ]
	FileInputStream file1;
	XSSFWorkbook workBook1;
	XSSFSheet sheet1;

	// this is for accessing file 2 first row [ creating object ]
	FileInputStream file2;
	XSSFWorkbook workBook2;
	XSSFSheet sheet2;

	// field
	String path1;
	String path2;
	int sheetNo1;
	int sheetNo2;
	int key1;
	int key2;

	// systemField
	String SystemFileFolder;
	String SystemFilePath1;
	String SystemFilePath2;
	String SystemFileName1;
	String SystemFileName2;
	int SystemFile1Sheet;
	int SystemFile2Sheet;
	String SystemSheetName;
	String SystemSheetName2;
	int Systemkey1;
	int Systemkey2;
	String SystemKeyName;
	String SystemKeyName2;
	String SystemFolderPath;

	// constructor called
	private Excel() {

		// setting title
		super("EXCEL COMPARATOR");

		// setting layout
		setLayout(new GridBagLayout());

		GridBagConstraints constraints = new GridBagConstraints();
		constraints.anchor = GridBagConstraints.WEST;
		constraints.insets = new Insets(10, 10, 10, 10);

		// getting data from configuration
		String projectPath = System.getProperty("user.dir");

		try {
			FileReader reader = new FileReader(projectPath + "\\config.txt");
			Properties properties = new Properties();
			properties.load(reader);

			SystemFileFolder = properties.getProperty("fileFolder");
			SystemFolderPath = properties.getProperty("folderPath");

			SystemFilePath1 = properties.getProperty("file1");
			SystemFilePath2 = properties.getProperty("file2");

			if (SystemFilePath1.equals("") && SystemFilePath2.equals("")) {
				SystemFilePath1 = null;
				SystemFilePath2 = null;
			}

			if (SystemFilePath1 != null && SystemFilePath2 != null) {

				try {
					SystemFile1Sheet = Integer.parseInt(properties.getProperty("sheetNo1"));
					Systemkey1 = Integer.parseInt(properties.getProperty("key1"));
					selectSheet1Drop.removeAllItems();
					file1 = new FileInputStream(SystemFilePath1);
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

					if (sheet1.getRow(0) == null) {
						JOptionPane.showMessageDialog(Excel.this, "Header Row is empty", "Excel",
								JOptionPane.ERROR_MESSAGE);
						SystemFilePath1 = null;
					} else {
						int column = sheet1.getRow(0).getLastCellNum();
						XSSFRow row = sheet1.getRow(0);
						for (int c = 0; c < column; c++) {
							if (row.getCell(c) == null) {
								headerDrop.addItem("");
							} else {
								headerDrop.addItem("" + row.getCell(c));
							}
						}
					}

					headerDrop.setSelectedIndex(Systemkey1);

				} catch (NumberFormatException e) {

					try {
						selectSheet1Drop.removeAllItems();
						file1 = new FileInputStream(SystemFilePath1);
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
							JOptionPane.showMessageDialog(Excel.this, "Excel file 1 is Empty", "Excel",
									JOptionPane.ERROR_MESSAGE);
							SystemFilePath1 = null;
						} else {
							int column = sheet1.getRow(0).getLastCellNum();
							XSSFRow row = sheet1.getRow(0);
							for (int c = 0; c < column; c++) {
								if (row.getCell(c) == null) {
									headerDrop.addItem("");
								} else {
									headerDrop.addItem("" + row.getCell(c));
								}
							} // for

						}

					} catch (FileNotFoundException ee) {
						JOptionPane.showMessageDialog(Excel.this,
								"Excel File 1 is already openned somewhere, please close the excel file "
										+ "/ or File Path is wrong",
								"Excel", JOptionPane.WARNING_MESSAGE);
						SystemFilePath1 = "";
						displayFileName1.setText("");
					}
				} catch (FileNotFoundException fee) {

				}

				try {
					SystemFile2Sheet = Integer.parseInt(properties.getProperty("sheetNo2"));
					Systemkey2 = Integer.parseInt(properties.getProperty("key2"));

					selectSheet2Drop2.removeAllItems();
					file2 = new FileInputStream(SystemFilePath2);
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
						JOptionPane.showMessageDialog(Excel.this, "Excel file 2 is Empty", "Excel",
								JOptionPane.ERROR_MESSAGE);
						SystemFilePath2 = null;
					} else {

						int column = sheet2.getRow(0).getLastCellNum();
						XSSFRow row = sheet2.getRow(0);

						for (int c = 0; c < column; c++) {
							if (row.getCell(c) == null) {
								headerDrop2.addItem("");
							} else {
								headerDrop2.addItem("" + row.getCell(c));
							}
						} // for
					}

					headerDrop2.setSelectedIndex(Systemkey2);

				} catch (NumberFormatException e) {

					try {
						selectSheet2Drop2.removeAllItems();
						file2 = new FileInputStream(SystemFilePath2);
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
							JOptionPane.showMessageDialog(Excel.this, "Excel file 2 is Empty", "Excel",
									JOptionPane.ERROR_MESSAGE);
							SystemFilePath2 = null;
						} else {

							int column = sheet2.getRow(0).getLastCellNum();
							XSSFRow row = sheet2.getRow(0);

							for (int c = 0; c < column; c++) {
								if (row.getCell(c) == null) {
									headerDrop2.addItem("");
								} else {
									headerDrop2.addItem("" + row.getCell(c));
								}
							} // for
						}

					} catch (FileNotFoundException ee) {
						JOptionPane.showMessageDialog(Excel.this,
								"Excel File 2 is already openned somewhere, please close the excel file "
										+ "/ or File Path is wrong",
								"Excel", JOptionPane.WARNING_MESSAGE);
						SystemFilePath2 = "";
						displayFileName1.setText("");
					}
				} catch (FileNotFoundException e) {
				}

				try {
					File filePath1 = new File(SystemFilePath1);
					fileName1 = filePath1.getName();
				} catch (Exception e) {
				}

				try {
					File filePath2 = new File(SystemFilePath2);
					fileName2 = filePath2.getName();
				} catch (Exception e) {
				}

			}

		} catch (FileNotFoundException fe) {
			JOptionPane.showMessageDialog(Excel.this, "configuration file not found", "Excel",
					JOptionPane.ERROR_MESSAGE);
		} catch (IOException e1) {
			e1.printStackTrace();
		}

		// adding child components at particular coordinates
		constraints.gridy = 0;
		constraints.gridx = 0;
		add(labelFILE1, constraints);

		constraints.gridy = 0;
		constraints.gridx = 1;
		add(buttonFile1, constraints);

		constraints.gridy = 2;
		constraints.gridx = 0;
		add(selectSheet1, constraints);

		constraints.gridy = 2;
		constraints.gridx = 1;
		add(selectSheet1Drop, constraints);

		constraints.gridy = 3;
		constraints.gridx = 0;
		add(selectSheet2, constraints);

		constraints.gridy = 3;
		constraints.gridx = 1;
		add(selectSheet2Drop2, constraints);

		// adding action listener to buttons
		buttonFile1.addActionListener((e) -> {

			// if button is pressed then pop-up window will appear
			if (e.getSource() == buttonFile1) {

				// remove predefined when clicking on button1
				SystemFilePath1 = null;

				selectSheet1Drop.removeAllItems();
				headerDrop.removeAllItems();

				displayFileName1.setText("");

				JFileChooser fileChooser = new JFileChooser();

				// setting this for only .xlsx
				FileNameExtensionFilter fnef = new FileNameExtensionFilter("Excel file (.xlsx)", "xlsx", "xlsb", "xls");
				fileChooser.setFileFilter(fnef);

				if (SystemFileFolder != null) {
					fileChooser.setCurrentDirectory(new File(SystemFileFolder));
				}

				Action details = fileChooser.getActionMap().get("viewTypeDetails");
				details.actionPerformed(null);

				int response = fileChooser.showOpenDialog(null);

				if (response == JFileChooser.APPROVE_OPTION) {

					File filePath1 = fileChooser.getSelectedFile();

					fileName1 = filePath1.getName();

					if (filePath1.getName().length() < 12) {
						displayFileName1.setText(filePath1.getName());
					} else {
						displayFileName1.setText(filePath1.getName().substring(0, 12));
					}

					String s = fileChooser.getSelectedFile().getAbsolutePath();

					path1 = s;

					if (path2 != null) {

						if (path1.equals(path2)) {
							JOptionPane.showMessageDialog(Excel.this,
									"Both File 1 and File 2 are Same Select other file", "Excel",
									JOptionPane.ERROR_MESSAGE);
							displayFileName1.setText("");
						}
					} else {

						try {

							selectSheet1Drop.removeAllItems();
							file1 = new FileInputStream(path1);
							workBook1 = new XSSFWorkbook(file1);

							int numberOfSheet1 = workBook1.getNumberOfSheets();

							for (int i = 0; i < numberOfSheet1; i++) {
								selectSheet1Drop.addItem(workBook1.getSheetName(i));
							}

						} catch (XLSBUnsupportedException ewww) {
							JOptionPane.showMessageDialog(Excel.this, "File format is not supported", "Excel",
									JOptionPane.WARNING_MESSAGE);
							displayFileName1.setText("");
							SystemFilePath1 = null;
							path1 = null;
						} catch (FileNotFoundException ee) {
							JOptionPane.showMessageDialog(Excel.this,
									"File is already open, please close the " + "file !", "Excel",
									JOptionPane.WARNING_MESSAGE);
							path1 = "";
							displayFileName1.setText("");
						} catch (IOException e1) {
						}
					}
				}
			}
		});

		selectSheet1Drop.setPrototypeDisplayValue("XXXXXXXXXXXXXXXXXXX");
		selectSheet1Drop.setMaximumRowCount(5);

		selectSheet1Drop.addActionListener((e) -> {

			if (e.getSource() == selectSheet1Drop) {

				headerDrop.removeAllItems();

				int selectedSheet1 = selectSheet1Drop.getSelectedIndex();
				SystemFile1Sheet = selectedSheet1;
				SystemSheetName = (String) selectSheet1Drop.getSelectedItem();

				try {
					sheet1 = workBook1.getSheetAt(selectedSheet1);
				} catch (IllegalArgumentException dd) {
				}

				sheetNo1 = selectedSheet1;
				sheetName1 = sheet1.getSheetName();

				if (sheet1.getRow(0) == null) {
					JOptionPane.showMessageDialog(Excel.this, "Excel file 1 is Empty", "Excel",
							JOptionPane.ERROR_MESSAGE);
					path1 = null;
				} else {
					int column = sheet1.getRow(0).getLastCellNum();

					XSSFRow row = sheet1.getRow(0);
					for (int c = 0; c < column; c++) {
						if (row.getCell(c) == null) {
							headerDrop.addItem("");
						} else {
							headerDrop.addItem("" + row.getCell(c));
						}
					} // for

				}
			}
		});

		// for predefined display
		try {
			File filePath1 = new File(SystemFilePath1);
			if (filePath1.getName().length() < 12) {
				displayFileName1.setText(filePath1.getName());
			} else {
				displayFileName1.setText(filePath1.getName().substring(0, 12));
			}
		} catch (Exception e) {
		}

		try {
			File filePath2 = new File(SystemFilePath2);

			if (filePath2.getName().length() < 12) {
				displayFileName2.setText(filePath2.getName());
			} else {
				displayFileName2.setText(filePath2.getName().substring(0, 12));
			}
		} catch (Exception e) {
		}

		try {
			File folderPathh = new File(SystemFolderPath);

			if (folderPathh.getName().length() < 12) {
				displayOutputFolder.setText(folderPathh.getName());
			} else {
				displayOutputFolder.setText(folderPathh.getName().substring(0, 12));
			}
		} catch (Exception e) {
		}

		constraints.gridy = 1;
		constraints.gridx = 1;
		add(buttonFile2, constraints);

		constraints.gridy = 1;
		constraints.gridx = 2;
		add(displayFileName2, constraints);

		buttonFile2.addActionListener((e) -> {
			if (e.getSource() == buttonFile2) {

				selectSheet2Drop2.removeAllItems();
				headerDrop2.removeAllItems();

				displayFileName2.setText("");

				JFileChooser fileChooser = new JFileChooser();

				FileNameExtensionFilter fnef = new FileNameExtensionFilter("Excel file (.xlsx)", "xlsx", "xlsb", "xls");
				fileChooser.setFileFilter(fnef);

				if (SystemFileFolder != null) {
					fileChooser.setCurrentDirectory(new File(SystemFileFolder));
				}

				Action details = fileChooser.getActionMap().get("viewTypeDetails");
				details.actionPerformed(null);

				int response = fileChooser.showOpenDialog(null);

				if (response == JFileChooser.APPROVE_OPTION) {
					File file11 = fileChooser.getSelectedFile();

					fileName2 = file11.getName();
					if (file11.getName().length() < 12) {
						displayFileName2.setText(file11.getName());
					} else {
						displayFileName2.setText(file11.getName().substring(0, 12));
					}

					String s = fileChooser.getSelectedFile().getAbsolutePath();
					path2 = s;

					if (path1 != null) {
						if (path1.equals(path2)) {
							JOptionPane.showMessageDialog(Excel.this,
									"Both File 1 and File 2 are Same Select other file", "File",
									JOptionPane.ERROR_MESSAGE);
							displayFileName2.setText("");
						} else {

							try {
								selectSheet2Drop2.removeAllItems();
								file2 = new FileInputStream(path2);
								workBook2 = new XSSFWorkbook(file2);

								int numberOfSheet2 = workBook2.getNumberOfSheets();

								for (int i = 0; i < numberOfSheet2; i++) {
									selectSheet2Drop2.addItem(workBook2.getSheetName(i));
								}

							} catch (XLSBUnsupportedException ewww) {
								JOptionPane.showMessageDialog(Excel.this, "File format is not supported", "Excel",
										JOptionPane.WARNING_MESSAGE);
								displayFileName2.setText("");
								SystemFilePath2 = null;
								path2 = null;

							} catch (FileNotFoundException ee) {
								JOptionPane.showMessageDialog(Excel.this,
										"File is already open, please close the " + "file !", "Excel",
										JOptionPane.WARNING_MESSAGE);
								path2 = "";
								displayFileName2.setText("");
							} catch (IOException e1) {
								e1.printStackTrace();
							}
						}
					}
				}
			}
		});

		selectSheet2Drop2.setPrototypeDisplayValue("XXXXXXXXXXXXXXXXXXX");
		selectSheet2Drop2.setMaximumRowCount(5);

		selectSheet2Drop2.addActionListener((e) -> {
			if (e.getSource() == selectSheet2Drop2) {

				headerDrop2.removeAllItems();
				int selectedSheet2 = selectSheet2Drop2.getSelectedIndex();
				SystemFile2Sheet = selectedSheet2;
				SystemSheetName2 = (String) selectSheet2Drop2.getSelectedItem();

				try {
					sheet2 = workBook2.getSheetAt(selectedSheet2);
				} catch (IllegalArgumentException dd) {

				}

				sheetNo2 = selectedSheet2;
				sheetName2 = sheet2.getSheetName();

				if (sheet2.getRow(0) == null) {
					JOptionPane.showMessageDialog(Excel.this, "Excel file 2 is Empty", "Excel",
							JOptionPane.ERROR_MESSAGE);
					path2 = null;
				} else {
					int column = sheet2.getRow(0).getLastCellNum();
					XSSFRow row = sheet2.getRow(0);

					for (int c = 0; c < column; c++) {
						if (row.getCell(c) == null) {
							headerDrop2.addItem("");
						} else {
							headerDrop2.addItem("" + row.getCell(c));
						}
					} // for
				}
			}
		});

		constraints.gridy = 0;
		constraints.gridx = 2;
		add(displayFileName1, constraints);
		
		headerDrop.setPrototypeDisplayValue("XXXXXXXXXXXXXXXXXXX");
		headerDrop.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {
				if (e.getSource() == headerDrop) {
					
					key1 = headerDrop.getSelectedIndex();
					Systemkey1 = key1;
					SystemKeyName = (String) headerDrop.getSelectedItem();
					keyName1 = (String) headerDrop.getSelectedItem();
				}
			}
		});

		// And limit the maximum number of items displayed in the drop-down list:
		headerDrop.setMaximumRowCount(5);

		constraints.gridy = 1;
		constraints.gridx = 0;
		add(labelFILE2, constraints);

		headerDrop2.setPrototypeDisplayValue("XXXXXXXXXXXXXXXXXXX");
		headerDrop2.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {
				if (e.getSource() == headerDrop2) {
					
					key2 = headerDrop2.getSelectedIndex();
					Systemkey2 = key2;
					SystemKeyName2 = (String) headerDrop.getSelectedItem();
					keyName2 = (String) headerDrop2.getSelectedItem();
				}
			}
		});

		headerDrop2.setMaximumRowCount(5);

		constraints.gridy = 5;
		constraints.gridx = 0;
		add(labelKEYFILE1, constraints);

		constraints.gridy = 5;
		constraints.gridx = 1;
		add(headerDrop, constraints);

		constraints.gridy = 6;
		constraints.gridx = 0;
		add(labelKEYFILE2, constraints);

		constraints.gridy = 6;
		constraints.gridx = 1;
		add(headerDrop2, constraints);

		constraints.gridx = 0;
		constraints.gridy = 7;
		add(outputFolder, constraints);

		constraints.gridx = 1;
		constraints.gridy = 7;
		add(buttonOutput, constraints);

		constraints.gridx = 2;
		constraints.gridy = 7;
		add(displayOutputFolder, constraints);

		buttonOutput.addActionListener(new ActionListener() {

			public void actionPerformed(ActionEvent e) {
				if (e.getSource() == buttonOutput) {

					JFileChooser fileChooser = new JFileChooser();
					fileChooser.setFileSelectionMode(JFileChooser.DIRECTORIES_ONLY);

					if (SystemFileFolder != null) {
						fileChooser.setCurrentDirectory(new File(SystemFolderPath));
					}
					
					int response = fileChooser.showOpenDialog(Excel.this);

					if (response == JFileChooser.APPROVE_OPTION) {
						file = new File(fileChooser.getSelectedFile().getAbsolutePath());
						File file2 = fileChooser.getSelectedFile();
						if (file2.getName().length() < 12) {
							displayOutputFolder.setText(file2.getName());
						} else {
							displayOutputFolder.setText(file2.getName().substring(0, 12));
						}

						String s = fileChooser.getSelectedFile().getAbsolutePath();
						SystemFolderPath = s;
					} else {
						displayOutputFolder.setText("");
					}
				}
			}
		});

		constraints.gridx = 0;
		constraints.gridy = 8;

		add(buttonENTER, constraints);

		buttonENTER.setBackground(Color.cyan);

		buttonENTER.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent event) {

				if (SystemFilePath1 != null && SystemFilePath2 != null && SystemFolderPath != null
						&& SystemFile1Sheet >= 0 && SystemFile2Sheet >= 0 && Systemkey1 >= 0 && Systemkey2 >= 0) {

					int e = fetchExcel(SystemFilePath1, SystemFilePath2, SystemFile1Sheet, SystemFile2Sheet, Systemkey1,
							Systemkey2, fileName1, fileName2, SystemFolderPath);

					if (e <= 1) {
						JOptionPane.showMessageDialog(Excel.this, "Excel created", "Excel", JOptionPane.PLAIN_MESSAGE);

						file = new File(SystemFolderPath);

						int ii = JOptionPane.showConfirmDialog(null,
								"We Have to close this window in order to open newly generated Excel, Because these are already open or are in use by javaw.exe Or if have to get more excels then click on No",
								"Exit?", JOptionPane.YES_NO_OPTION);
						if (ii == 1) {
							// do nothing
						}
						if (ii == 0) {
							try {
								desktop.open(file);
							} catch (IOException eeee) {
								eeee.printStackTrace();
							}
							System.exit(0);
						}
					}

				} else {

					if (path1 == null && path2 != null) {
						JOptionPane.showMessageDialog(Excel.this, "Enter File1", "File", JOptionPane.ERROR_MESSAGE);
					} else if (path1 == null && path2 == null) {
						JOptionPane.showMessageDialog(Excel.this, "Enter Files", "File", JOptionPane.ERROR_MESSAGE);
					} else if (path1 != null && path2 == null) {
						JOptionPane.showMessageDialog(Excel.this, "Enter File2", "File", JOptionPane.ERROR_MESSAGE);
					} else if (path1 != null && path2 != null && SystemFolderPath == null) {
						JOptionPane.showMessageDialog(Excel.this, "Select Folder", "Folder", JOptionPane.ERROR_MESSAGE);
					} else if (path1 == null && path2 == null && SystemFolderPath == null) {
						JOptionPane.showMessageDialog(Excel.this, "Select Files and Folder", "Folder",
								JOptionPane.ERROR_MESSAGE);
					} else if (path1.equals(path2)) {
						JOptionPane.showMessageDialog(Excel.this, "Both File 1 and File 2 are Same Select other file",
								"File", JOptionPane.ERROR_MESSAGE);
						path2 = null;
						displayFileName2.setText("");
						headerDrop2.removeAllItems();
					} else if (path1 != null && path2 != null && SystemFolderPath != null) {

						int e = fetchExcel(path1, path2, sheetNo1, sheetNo2, key1, key2, fileName1, fileName2,
								sheetName1, sheetName2, keyName1, keyName2, SystemFolderPath);

						if (e <= 1) {
							JOptionPane.showMessageDialog(Excel.this, "Excel created", "Excel",
									JOptionPane.PLAIN_MESSAGE);

							file = new File(SystemFolderPath);

							int ii = JOptionPane.showConfirmDialog(null,
									"We Have to close this window in order to open newly generated Excel, Because these are already open or are in use by javaw.exe Or if have to get more excels then click on No",
									"Exit?", JOptionPane.YES_NO_OPTION);
							if (ii == 1) {
							}
							if (ii == 0) {
								try {
									desktop.open(file);
								} catch (IOException eeee) {
									eeee.printStackTrace();
								}
								System.exit(0);
							}
						}
					}
				}
			}
		});

		constraints.gridx = 1;
		constraints.gridy = 8;

		add(buttonDUPLICATE, constraints);

		buttonDUPLICATE.setBackground(Color.cyan);

		buttonDUPLICATE.addActionListener((e) -> {

			if (SystemFilePath1 != null && SystemFilePath2 != null && SystemFolderPath != null && SystemFile1Sheet >= 0
					&& SystemFile2Sheet >= 0 && Systemkey1 >= 0 && Systemkey2 >= 0) {

				int eq = duplicateExcel(SystemFilePath1, SystemFilePath2, SystemFile1Sheet, SystemFile2Sheet,
						Systemkey1, Systemkey2, fileName1, fileName2, SystemFolderPath);

				if (eq <= 1) {
					JOptionPane.showMessageDialog(Excel.this, "Excel created", "Excel", JOptionPane.PLAIN_MESSAGE);

					file = new File(SystemFolderPath);

					int ii = JOptionPane.showConfirmDialog(null,
							"We Have to close this window in order to open newly generated Excel, Because these are already open or are in use by javaw.exe Or if have to get more excels then click on No",
							"Exit?", JOptionPane.YES_NO_OPTION);
					if (ii == 1) {
						// do nothing
					}
					if (ii == 0) {
						try {
							desktop.open(file);
						} catch (IOException eeee) {
							eeee.printStackTrace();
						}
						System.exit(0);
					}
				}

			} else {

				if (path1 == null && path2 != null) {
					JOptionPane.showMessageDialog(Excel.this, "Enter File1", "File", JOptionPane.ERROR_MESSAGE);
				} else if (path1 == null && path2 == null) {
					JOptionPane.showMessageDialog(Excel.this, "Enter Files", "File", JOptionPane.ERROR_MESSAGE);
				} else if (path1 != null && path2 == null) {
					JOptionPane.showMessageDialog(Excel.this, "Enter File2", "File", JOptionPane.ERROR_MESSAGE);
				} else if (path1 != null && path2 != null && SystemFolderPath == null) {
					JOptionPane.showMessageDialog(Excel.this, "Select Folder", "Folder", JOptionPane.ERROR_MESSAGE);
				} else if (path1 == null && path2 == null && SystemFolderPath == null) {
					JOptionPane.showMessageDialog(Excel.this, "Select Files and Folder", "Folder",
							JOptionPane.ERROR_MESSAGE);
				} else if (path1.equals(path2)) {
					JOptionPane.showMessageDialog(Excel.this, "Both File 1 and File 2 are Same Select other file",
							"File", JOptionPane.ERROR_MESSAGE);
					path2 = null;
					displayFileName2.setText("");
					headerDrop2.removeAllItems();
				} else if ((path1 != null && path2 != null && SystemFolderPath != null)) {

					int ee = duplicateExcel(path1, path2, sheetNo1, sheetNo2, key1, key2, fileName1, fileName2,
							sheetName1, sheetName2, keyName1, keyName2, SystemFolderPath);

					if (ee <= 1) {
						JOptionPane.showMessageDialog(Excel.this, "Excel created", "Excel", JOptionPane.PLAIN_MESSAGE);

						file = new File(SystemFolderPath);

						int ii = JOptionPane.showConfirmDialog(null,
								"We Have to close this window in order to open newly generated Excel, Because these are already open or are in use by javaw.exe Or if have to get more excels then click on No",
								"Exit?", JOptionPane.YES_NO_OPTION);
						if (ii == 1) {
							// do nothing
						}
						if (ii == 0) {
							try {
								desktop.open(file);
							} catch (IOException eeee) {
								eeee.printStackTrace();
							}
							System.exit(0);
						}

					}
				}
			}
		});

		constraints.gridx = 2;
		constraints.gridy = 8;
		add(buttonClear, constraints);

		buttonClear.setBackground(Color.red);

		buttonClear.addActionListener(new ActionListener() {

			public void actionPerformed(ActionEvent e) {

				path1 = null;
				path2 = null;

				selectSheet1Drop.removeAllItems();
				selectSheet2Drop2.removeAllItems();
				headerDrop.removeAllItems();
				headerDrop2.removeAllItems();

				SystemFilePath1 = null;
				SystemFilePath2 = null;
				SystemFolderPath = null;

				displayFileName1.setText("");
				displayFileName2.setText("");
				displayOutputFolder.setText("");

			}

		});

		pack();
		setResizable(false);
		setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
		setLocationRelativeTo(null);

	}

// ----------------------------------------------------------------------------------------

}

public class ExcelTask {

	public static void main(String[] args) {

		try {
			UIManager.setLookAndFeel(UIManager.getSystemLookAndFeelClassName());
		} catch (Exception e) {
		}

		SwingUtilities.invokeLater(new Runnable() {
			public void run() {
				Excel.getInstance().setVisible(true);
			}
		});
	}
}
