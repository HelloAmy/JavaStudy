package zjr.amy.excel;

import java.io.BufferedInputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Date;
import java.util.List;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;

/**
 * Excel帮助类
 * 
 * @author zhujinrong
 * 
 */
public class ExcelHelper {

	/**
	 * 读取Excel内容
	 * 
	 * @param fileName
	 *            文件名和路径
	 * @param ignoreRows
	 *            忽略读取的行数
	 * @return 数据
	 * @throws FileNotFoundException
	 *             找不到文件
	 * @throws IOException
	 *             输入输出异常
	 */
	public static String[][] getData(String fileName, int ignoreRows)
			throws FileNotFoundException, IOException {
		File file = new File(fileName);
		List<String[]> result = new ArrayList<String[]>();
		int rowSize = 0;
		BufferedInputStream in = new BufferedInputStream(new FileInputStream(
				file));
		POIFSFileSystem fs = new POIFSFileSystem(in);
		HSSFWorkbook workbook = new HSSFWorkbook(fs);
		HSSFCell cell = null;
		for (int sheetIndex = 0; sheetIndex < workbook.getNumberOfSheets(); sheetIndex++) {

			// 可能有多个sheet
			HSSFSheet sheet = workbook.getSheetAt(sheetIndex);
			for (int rowIndex = ignoreRows; rowIndex < sheet.getLastRowNum(); rowIndex++) {
				HSSFRow row = sheet.getRow(rowIndex);
				if (row == null) {
					continue;
				}

				// 获取行数
				int tempRowSize = row.getLastCellNum() + 1;
				if (tempRowSize > rowSize) {
					rowSize = tempRowSize;
				}

				String[] values = new String[rowSize];
				Arrays.fill(values, "");
				boolean hasValue = false;
				for (int collumnIndex = 0; collumnIndex <= row.getLastCellNum(); collumnIndex++) {
					String value = "";
					cell = row.getCell(collumnIndex);
					if (cell != null) {
						switch (cell.getCellType()) {
						case HSSFCell.CELL_TYPE_STRING:
							value = cell.getStringCellValue();
							break;
						case HSSFCell.CELL_TYPE_NUMERIC:
							if (HSSFDateUtil.isCellDateFormatted(cell)) {
								Date date = cell.getDateCellValue();
								if (date != null) {
									value = new SimpleDateFormat("yyyy-MM-dd")
											.format(date);
								} else {
									value = "";
								}
							} else {
								double d = cell.getNumericCellValue();
								value = d + "";
							}
							break;
						case HSSFCell.CELL_TYPE_FORMULA:
							if (!cell.getStringCellValue().equals("")) {
								value = cell.getStringCellValue();
							} else {
								value = cell.getNumericCellValue() + "";
							}
							break;
						case HSSFCell.CELL_TYPE_BLANK:
							break;
						case HSSFCell.CELL_TYPE_ERROR:
							break;
						case HSSFCell.CELL_TYPE_BOOLEAN:
							value = (cell.getBooleanCellValue() == true ? "Y"
									: "N");
							break;
						default:
							value = "";
						}
						if (collumnIndex == 0 && value.trim().equals("")) {
							break;
						}
						values[collumnIndex] = value.trim();
						hasValue = true;
					}
				}
				if (hasValue) {
					result.add(values);
				}
			}
		}

		in.close();
		String[][] returnArray = new String[result.size()][rowSize];
		for (int i = 0; i < returnArray.length; i++) {
			returnArray[i] = (String[]) result.get(i);
		}

		return returnArray;
	}

	/**
	 * 写Excel数据
	 * 
	 * @param fileName
	 *            文件名
	 * @param data
	 *            数据
	 * @param sheetName
	 *            sheet名字
	 * @throws IOException
	 *             输入输出异常
	 */
	public static void setData(String fileName, String[][] data,
			String sheetName) throws IOException {
		FileOutputStream fos = new FileOutputStream(fileName);
		HSSFWorkbook wb = new HSSFWorkbook();
		HSSFSheet s = wb.createSheet();
		wb.setSheetName(0, sheetName);
		if (data != null && data.length > 0) {
			for (int i = 0; i < data.length; i++) {
				// 创建第i行
				HSSFRow row = s.createRow(i);
				for (int j = 0; j < data[i].length; j++) {
					// 第i行第j列的数据
					HSSFCell cell = row.createCell(j);
					cell.setCellValue(data[i][j]);
				}
			}
		}

		// 写入工作薄
		wb.write(fos);

		// 关闭输出流
		fos.close();
	}
}
