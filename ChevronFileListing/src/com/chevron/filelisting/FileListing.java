package com.chevron.filelisting;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.StringReader;
import java.util.ArrayList;
import java.util.Collection;
import java.util.Properties;
import java.util.Scanner;

import javax.swing.JOptionPane;

import org.apache.poi.common.usermodel.HyperlinkType;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Hyperlink;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.util.CellAddress;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class FileListing {
	public FileListing() {
	}

	private int rowAkhir(XSSFSheet sheet, CellAddress caPath) {
		int row = 0;
		caPath = new CellAddress(caPath.toString());
		if (caPath.getRow() != sheet.getLastRowNum() + 1) {
			for (int x = caPath.getRow(); x < sheet.getLastRowNum() + 1; x++) {
				XSSFRow rowCount2 = sheet.getRow(x);
				XSSFCell cell = rowCount2.getCell(caPath.getColumn());
				if (rowCount2 != null && cell == null) {
					row = x - 1;
					break;

				} else if (rowCount2 != null && cell != null) {
					if (cell.toString().isEmpty()) {
						row = x - 1;
						break;
					} else {
						row = sheet.getLastRowNum();
					}
				}
			}
		} else {
			row = sheet.getLastRowNum();
		}
		return row;
	}

	private String getFileExtension(File file) {
		String name = file.getName();
		return name.substring(name.lastIndexOf(".") + 1);
	}
	
	//cari file. note: hidden juga di copy
	private void addTree(File file, Collection<File> all) {
		File[] children = file.listFiles();
		if (children != null) {
			for (File child : children) {
				if ((child.isFile())) {
					all.add(child);
					addTree(child, all);
				} else if (child.isDirectory()) {
					addTree(child, all);
				}
			}
		}
	}

	private void generate(File source, File dest, File sourceFolder, File path,
			File name, File type) throws IOException {
		Collection<File> all = new ArrayList<File>();
		addTree(sourceFolder, all);
		FileInputStream inputStream = new FileInputStream(source);
		XSSFWorkbook workbook = new XSSFWorkbook(inputStream);
		workbook.setForceFormulaRecalculation(true);
		CreationHelper createHelper = workbook.getCreationHelper();
		Hyperlink link = createHelper.createHyperlink(HyperlinkType.FILE);
		Font hlink_font = workbook.createFont();
		hlink_font.setUnderline((byte) 1);
		hlink_font.setColor(IndexedColors.BLUE_GREY.getIndex());
		XSSFSheet sheet = workbook.getSheetAt(0);
		XSSFCellStyle style = workbook.createCellStyle();
		XSSFCellStyle styleURI = workbook.createCellStyle();
		style.setBorderTop(BorderStyle.THIN);
		style.setBorderBottom(BorderStyle.THIN);
		style.setBorderLeft(BorderStyle.THIN);
		style.setBorderRight(BorderStyle.THIN);
		styleURI.setBorderTop(BorderStyle.THIN);
		styleURI.setBorderBottom(BorderStyle.THIN);
		styleURI.setBorderLeft(BorderStyle.THIN);
		styleURI.setBorderRight(BorderStyle.THIN);
		styleURI.setFont(hlink_font);
		XSSFRow row = null;
		XSSFCell cell = null;
		CellAddress caPath = new CellAddress(path.toString());
		CellAddress caName = new CellAddress(name.toString());
		CellAddress caType = new CellAddress(type.toString());

		int lastCol = sheet.getRow(caPath.getRow() - 1).getLastCellNum();

		int rowCount = rowAkhir(sheet, caPath);
		for (File elem : all) {
			rowCount++;

			if (sheet.getRow(rowCount) != null) {
				row = sheet.getRow(rowCount);
				if (row.getCell(caPath.getColumn()) == null) {
					cell = row.createCell(caPath.getColumn());
					cell.setCellStyle(style);
				} else {
					cell = row.getCell(caPath.getColumn());
				}
				link = createHelper.createHyperlink(HyperlinkType.FILE);
				cell.setHyperlink(link);
				cell.setCellValue(elem.getAbsolutePath());
				link.setAddress(elem.toURI().toString());
				cell.setCellStyle(styleURI);

				if (row.getCell(caName.getColumn()) == null) {
					cell = row.createCell(caName.getColumn());
					cell.setCellStyle(style);
				} else {
					cell = row.getCell(caName.getColumn());
				}
				cell.setCellValue(elem.getName());

				if (row.getCell(caType.getColumn()) == null) {
					cell = row.createCell(caType.getColumn());
					cell.setCellStyle(style);
				} else {
					cell = row.getCell(caType.getColumn());
				}
				cell.setCellValue(getFileExtension(elem.getAbsoluteFile()));
			}

			if (sheet.getRow(rowCount) == null) {
				row = sheet.createRow(rowCount);
				cell = row.createCell(caPath.getColumn());
				link = createHelper.createHyperlink(HyperlinkType.FILE);
				cell.setHyperlink(link);
				cell.setCellStyle(styleURI);
				cell.setCellValue(elem.getAbsolutePath());
				link.setAddress(elem.toURI().toString());

				cell = row.createCell(caName.getColumn());
				cell.setCellStyle(style);
				cell.setCellValue(elem.getName());

				cell = row.createCell(caType.getColumn());
				cell.setCellStyle(style);
				cell.setCellValue(getFileExtension(elem.getAbsoluteFile()));
				int lastCol2 = row.getLastCellNum();
				for (int x = lastCol2; x < lastCol; x++) {
					cell = row.createCell(x);
					cell.setCellStyle(style);
				}
			}
		}

		inputStream.close();
		FileOutputStream outputStream = new FileOutputStream(dest);
		workbook.write(outputStream);
		workbook.close();
		outputStream.close();
	}

	public static void main(String[] args) throws IOException,
	java.text.ParseException {
		FileListing fl = new FileListing();
		File destExcel = null;
		File config = new File(System.getProperty("user.dir")
				+ "\\listing-config.properties");
		Properties prop = new Properties();

		Scanner sc = null;
		try {
			String str = "";
			if (config.exists()) {
				sc = new Scanner(config);
				while (sc.hasNextLine()) {
					str = str + sc.nextLine() + "\n";
				}
			} else {
				if (config.toString().length() > 23) {
					config = new File(config.toString() + "\n");
				} else {
					config = new File(config.toString() + " .");
				}
				JOptionPane
				.showMessageDialog(
						null,
						config
						+ "does not correct or does not exist.\nPlease make sure your listing-config.properties exists.");
				System.exit(0);
			}

			prop.load(new StringReader(str.replace("\\", "\\\\")));

			File sourceFolder = new File(prop.getProperty("sourceFolder"));
			File sourceExcel = new File(prop.getProperty("sourceExcel"));
			destExcel = new File(prop.getProperty("destExcel"));
			File FilePath = new File(prop.getProperty("FilePath"));
			File FileName = new File(prop.getProperty("FileName"));
			File FileType = new File(prop.getProperty("FileType"));
			try {
				CellAddress caPath = new CellAddress(FilePath.toString());
				CellAddress caName = new CellAddress(FileName.toString());
				CellAddress caType = new CellAddress(FileType.toString());
				if ((caPath.getRow() + caType.getRow()) / 2 != caName.getRow()) {
					JOptionPane.showMessageDialog(null,
							FilePath.toString() + " & " + FileName.toString()
							+ " & " + FileType.toString()
							+ "\nShould be had the same row.");
					System.exit(0);
				}
			} catch (NumberFormatException e) {
				JOptionPane
				.showMessageDialog(null,
						"you must insert a valid cell address. Please check your configuration file");
				System.exit(0);
			}
			if ((sourceFolder.exists()) && (sourceFolder.isDirectory())
					&& (sourceExcel.exists())) {
				fl.generate(sourceExcel, destExcel, sourceFolder, FilePath,
						FileName, FileType);
				JOptionPane.showMessageDialog(null,
						"The excel file is generated successfully to\n"
								+ destExcel);
			} else if (!sourceFolder.exists()) {
				JOptionPane.showMessageDialog(null, "SourceFolder "
						+ sourceFolder + "\nincorrect or doesn't exist");
			} else if (sourceFolder.isFile()) {
				JOptionPane.showMessageDialog(null, "SourceFolder "
						+ sourceFolder + "\nis a file not a directory!");
			} else if (!sourceExcel.exists()) {
				if (sourceExcel.toString().length() > 37) {
					sourceExcel = new File("\n" + sourceExcel.toString());
				}

				JOptionPane
				.showMessageDialog(
						null,
						"The source excel file "
								+ sourceExcel
								+ " does not exist.\nPlease make sure "
								+ "the source excel file exists and same with your configuration");
			} else {
				JOptionPane
				.showMessageDialog(
						null,
						"Generating excel is failed because it is currently being used by another process path\n"
								+ destExcel
								+ "\n"
								+ "please check your privilage to modify this path \n"
								+ "and check the file is not opened by another proses");
			}
		} catch (IOException e) {
			JOptionPane.showMessageDialog(null,
					"Generating excel is failed because it is currently being used by another process path\n"
							+ destExcel);
		} catch (NullPointerException e) {
			JOptionPane.showMessageDialog(null,
					"You must set the required settings. Please check your configuration file. \n");
		} finally {
			if (sc != null) {
				sc.close();
			}
		}
	}
}
