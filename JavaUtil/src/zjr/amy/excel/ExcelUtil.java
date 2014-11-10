package zjr.amy.excel;

import java.io.IOException;
import java.io.OutputStream;
import java.lang.reflect.Field;
import java.lang.reflect.InvocationTargetException;
import java.lang.reflect.Method;
import java.text.SimpleDateFormat;
import java.util.Calendar;
import java.util.Collection;
import java.util.Date;
import java.util.Iterator;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFClientAnchor;
import org.apache.poi.hssf.usermodel.HSSFComment;
import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.hssf.usermodel.HSSFPatriarch;
import org.apache.poi.hssf.usermodel.HSSFRichTextString;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.HSSFColor;

public class ExcelUtil<T> {
	public void writeData(Collection<T> dataset, OutputStream out) {
		Date date = new Date();
		SimpleDateFormat sdf = new SimpleDateFormat("yyyyMMddHHmmss");
		String temp = sdf.format(date);
		this.writeData("测试文档" + temp, null, dataset, out, "yyyy-MM-dd");
	}

	public void writeData(String[] headers, Collection<T> dataset,
			OutputStream out) {
		Date date = new Date();
		SimpleDateFormat sdf = new SimpleDateFormat("yyyyMMddHHmmss");
		String temp = sdf.format(date);
		this.writeData("测试文档" + temp, headers, dataset, out, "yyyy-MM-dd");
	}

	public void writeData(String[] headers, Collection<T> dataset,
			OutputStream out, String pattern) {
		this.writeData("测试文档", headers, dataset, out, pattern);
	}

	public void writeData(String title, String[] headers,
			Collection<T> dataset, OutputStream out, String pattern) {

		// 声明一个工作薄
		HSSFWorkbook workbook = new HSSFWorkbook();

		// 生成一个表格
		HSSFSheet sheet = workbook.createSheet();

		// 设置表格默认宽度为15个字节
		sheet.setDefaultColumnWidth(15);

		// 生成和设置表头样式
		HSSFCellStyle headStyle = workbook.createCellStyle();
		headStyle.setFillForegroundColor(HSSFColor.SKY_BLUE.index);
		headStyle.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
		headStyle.setBorderBottom(HSSFCellStyle.BORDER_THIN);
		headStyle.setBorderLeft(HSSFCellStyle.BORDER_THIN);
		headStyle.setBorderRight(HSSFCellStyle.BORDER_THIN);
		headStyle.setBorderTop(HSSFCellStyle.BORDER_THIN);
		headStyle.setAlignment(HSSFCellStyle.ALIGN_CENTER);

		// 生成一个字体
		HSSFFont headFont = workbook.createFont();
		headFont.setColor(HSSFColor.VIOLET.index);
		headFont.setFontHeightInPoints((short) 12);
		headFont.setBoldweight(HSSFFont.BOLDWEIGHT_BOLD);

		// 把字体应用到当前的样式
		headStyle.setFont(headFont);

		// 生成另一种样式
		HSSFCellStyle style2 = workbook.createCellStyle();

		// 设置这些样式
		style2.setFillForegroundColor(HSSFColor.LIGHT_YELLOW.index);
		style2.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
		style2.setBorderBottom(HSSFCellStyle.BORDER_THIN);
		style2.setBorderLeft(HSSFCellStyle.BORDER_THIN);
		style2.setBorderRight(HSSFCellStyle.BORDER_THIN);
		style2.setBorderTop(HSSFCellStyle.BORDER_THIN);
		style2.setAlignment(HSSFCellStyle.ALIGN_CENTER);
		style2.setVerticalAlignment(HSSFCellStyle.VERTICAL_CENTER);

		// 生成另一种字体
		HSSFFont font2 = workbook.createFont();
		font2.setBoldweight(HSSFFont.BOLDWEIGHT_NORMAL);

		// 把字体应用到当前样式
		style2.setFont(font2);

		// 声明一个画图的顶级管理器
		HSSFPatriarch patriarch = sheet.createDrawingPatriarch();

		// 定义注释的大小和位置，详见文档
		HSSFComment comment = patriarch.createComment(new HSSFClientAnchor(0,
				0, 0, 0, (short) 4, 2, (short) 6, 5));
		comment.setString(new HSSFRichTextString("可以在POI中添加注释！"));
		// 设置注释作者,当鼠标移动到单元格上是可以在状态栏看到的
		comment.setAuthor("amy");

		// 创建标题列
		HSSFRow row = sheet.createRow(0);
		if (headers != null) {
			for (int i = 0; i < headers.length; i++) {
				HSSFCell cell = row.createCell(i);
				cell.setCellStyle(headStyle);
				HSSFRichTextString text = new HSSFRichTextString(headers[i]);
				cell.setCellValue(text);
			}
		}
		Iterator<T> it = dataset.iterator();
		int index = 0;
		while (it.hasNext()) {
			index++;
			row = sheet.createRow(index);
			T t = (T) it.next();
			// 利用反射，根据java本属性的先后顺序，动态调用getX()方法得到属性值
			Field[] fields = t.getClass().getDeclaredFields();
			for (int i = 0; i < fields.length; i++) {
				HSSFCell cell = row.createCell(i);
				cell.setCellStyle(style2);
				Field field = fields[i];
				String fieldName = field.getName();
				String getMethodName = "get"
						+ fieldName.substring(0, 1).toUpperCase()
						+ fieldName.substring(1);
				try {
					Class<? extends Object> tcls = t.getClass();
					Method getMethod = tcls.getMethod(getMethodName,
							new Class[] {});
					Object value = getMethod.invoke(t, new Object[] {});
					this.setCellValue(cell, value, pattern, patriarch);
				} catch (SecurityException e) {
					e.printStackTrace();
				} catch (NoSuchMethodException e) {
					// TODO Auto-generated catch block
					e.printStackTrace();
				} catch (IllegalAccessException e) {
					// TODO Auto-generated catch block
					e.printStackTrace();
				} catch (IllegalArgumentException e) {
					// TODO Auto-generated catch block
					e.printStackTrace();
				} catch (InvocationTargetException e) {
					// TODO Auto-generated catch block
					e.printStackTrace();
				}
			}
		}

		try {
			workbook.write(out);
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		try {
			out.close();
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
	}

	public void setCellValue(HSSFCell cell, Object value, String pattern,
			HSSFPatriarch patriarch) {
		// 判断值的类型后进行强制类型转换
		String textValue = null;
		if (value instanceof Integer) {
			int intValue = (Integer) value;
			cell.setCellValue(intValue);
		} else if (value instanceof Float) {
			float floatValue = (float) value;
			HSSFRichTextString txtValue = new HSSFRichTextString(
					String.valueOf(floatValue));
			cell.setCellValue(txtValue);
		} else if (value instanceof Double) {
			double dValue = (double) value;
			HSSFRichTextString txtValue = new HSSFRichTextString(
					String.valueOf(dValue));
			cell.setCellValue(txtValue);
		} else if (value instanceof Long) {
			long lValue = (long) value;
			HSSFRichTextString txtValue = new HSSFRichTextString(
					String.valueOf(lValue));
			cell.setCellValue(txtValue);
		} else if (value instanceof Boolean) {
			boolean bValue = (boolean) value;
			textValue = "否";
			if (bValue) {
				textValue = "是";
			}
			cell.setCellValue(textValue);
		} else if (value instanceof Date) {
			Date date = (Date) value;
			SimpleDateFormat sdf = new SimpleDateFormat(pattern);
			textValue = sdf.format(date);
			cell.setCellValue(textValue);
		} else if (value instanceof byte[]) {
			// 有图片时，设置行高为60px;
			cell.getRow().setHeightInPoints(60);

			// 设置图片所在列宽度为80px,注意这里单位的一个换算
			cell.getRow().getSheet()
					.setColumnWidth(cell.getColumnIndex(), (int) (35.7 * 80));

			byte[] bsValue = (byte[]) value;
			HSSFClientAnchor anchor = new HSSFClientAnchor(0, 0, 1023, 255,
					(short) 6, cell.getRowIndex(), (short) 6,
					cell.getRowIndex());
			anchor.setAnchorType(2);
			patriarch.createPicture(anchor, cell.getSheet().getWorkbook()
					.addPicture(bsValue, HSSFWorkbook.PICTURE_TYPE_JPEG));

		} else if (value instanceof Calendar) {
			Calendar cValue = (Calendar) value;
			int month = (cValue.get(Calendar.MONTH) + 1);
			String monthStr = "";
			if (month < 10) {
				monthStr = "0" + month;
			} else {
				monthStr = month + "";
			}
			String txtValue = cValue.get(Calendar.YEAR) + "/"
					+ monthStr + "/"
					+ cValue.get(Calendar.DAY_OF_MONTH);
			cell.setCellValue(txtValue);
		} else {
			// 其他数据都当做字符串简单处理
			textValue = value.toString();
			cell.setCellValue(textValue);
		}

	}

}
