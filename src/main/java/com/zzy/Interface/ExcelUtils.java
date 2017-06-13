package com.zzy.Interface;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.List;

import javax.xml.soap.Text;

import org.apache.commons.lang.ObjectUtils.Null;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.omg.PortableInterceptor.USER_EXCEPTION;

public class ExcelUtils {

	public static Object[][] getTableArray(String FilePath) throws Exception {
		String[][] tabArray = null;
		String[][] tabArray2 = null;

		// InputStream is = this.getClass().getResourceAsStream(FilePath);
		InputStream is = new FileInputStream(FilePath);
		XSSFWorkbook xssfWorkbook = new XSSFWorkbook(is);
		XSSFSheet xssfSheet = xssfWorkbook.getSheetAt(0);
		// Read the Sheet
		// for (int numSheet = 0; numSheet < xssfWorkbook.getNumberOfSheets();
		// numSheet++) {
		// if (xssfSheet == null) {
		// continue;
		// }
		tabArray = new String[xssfSheet.getLastRowNum() + 1][xssfSheet.getRow(0).getLastCellNum()];
		tabArray2 = new String[xssfSheet.getLastRowNum() + 1-1][xssfSheet.getRow(0).getLastCellNum()];

		for (int rowNum = 1; rowNum < xssfSheet.getLastRowNum() + 1; rowNum++) {
			XSSFRow xssfRow = xssfSheet.getRow(rowNum);
			if (xssfRow != null) {
				for (int rowCol = 0; rowCol < xssfSheet.getRow(0).getLastCellNum(); rowCol++) {
					tabArray[rowNum][rowCol] = ExcelUtils.getCellData(xssfSheet.getRow(rowNum).getCell(rowCol));
//					String text = ExcelUtils.getCellData(xssfSheet.getRow(rowNum).getCell(rowCol));
//					System.out.println(text + " ");
					tabArray2[rowNum-1][rowCol]=tabArray[rowNum][rowCol];
				}
//				System.out.println("");
			}
		}
//		System.err.println(tabArray2[0][0]);
//		System.err.println(tabArray2[xssfSheet.getLastRowNum() + 1-2][xssfSheet.getRow(0).getLastCellNum()-1]);
		// }

		return tabArray2;
	}

	/**
	 * 获取单元格数据内容为字符串类型的数据
	 * 
	 * @param cell
	 *            Excel单元格
	 * @return String 单元格数据内容
	 */
	public static String getCellData(Cell cell) {
		String value = "";
		switch (cell.getCellType()) {
		case Cell.CELL_TYPE_NUMERIC: // 数值型
			if (DateUtil.isCellDateFormatted(cell)) {
				// 如果是date类型则 ，获取该cell的date值
				value = DateUtil.getJavaDate(cell.getNumericCellValue()).toString();
			} else {// 纯数字
				value = String.valueOf(cell.getNumericCellValue());
			}
			break;
		/* 此行表示单元格的内容为string类型 */
		case Cell.CELL_TYPE_STRING: // 字符串型
			value = cell.getRichStringCellValue().toString();
			break;
		case Cell.CELL_TYPE_FORMULA:// 公式型
			// 读公式计算值
			value = String.valueOf(cell.getNumericCellValue());
			if (value.equals("NaN")) {// 如果获取的数据值为非法值,则转换为获取字符串
				value = cell.getRichStringCellValue().toString();
			}
			// cell.getCellFormula();读公式
			break;
		case Cell.CELL_TYPE_BOOLEAN:// 布尔
			value = " " + cell.getBooleanCellValue();
			break;
		/* 此行表示该单元格值为空 */
		case Cell.CELL_TYPE_BLANK: // 空值
			value = "";
			break;
		case Cell.CELL_TYPE_ERROR: // 故障
			value = "";
			break;
		default:
			value = cell.getRichStringCellValue().toString();
		}
		return value;
	}

	
	public static void writeExcel(String path,String line,String output,String error_code,String error_message,String result ) throws Exception{
		InputStream is = new FileInputStream(path);
		XSSFWorkbook xssfWorkbook = new XSSFWorkbook(is);
		XSSFSheet xssfSheet = xssfWorkbook.getSheetAt(0);
//		System.out.println("line :"+line);
		int lineNum = (int)Double.parseDouble(line);
//		System.out.println("lineNum : "+lineNum);
		xssfSheet.getRow(lineNum).getCell(8).setCellValue(output);
		xssfSheet.getRow(lineNum).getCell(9).setCellValue(error_code);
		xssfSheet.getRow(lineNum).getCell(10).setCellValue(error_message);
		xssfSheet.getRow(lineNum).getCell(11).setCellValue(result);
//		System.out.println("output : "+output);
		
        OutputStream stream = new FileOutputStream(path);  
        xssfWorkbook.write(stream);  
        xssfWorkbook.close();
        stream.close();
        is.close();
	}
	
	/*public static void main(String[] args) {
		try {
			System.out.println(new ExcelUtils()
					.getTableArray(System.getProperty("user.dir") + "\\lib\\InterfaceModel.xlsx"));
		} catch (Exception e) {
			// TODO 自动生成的 catch 块
			e.printStackTrace();
		}
	}*/
}
