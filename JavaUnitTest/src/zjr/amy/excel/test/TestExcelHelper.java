package zjr.amy.excel.test;

import java.io.FileNotFoundException;
import java.io.IOException;

import zjr.amy.excel.ExcelHelper;

public class TestExcelHelper {

	/**
	 * @param args
	 */
	public static void main(String[] args) {
		// TODO Auto-generated method stub
		TestGetData();
	}

	/**
	 * 测试获取数据
	 */
	public static void TestGetData() {
		try {
			String[][] array = ExcelHelper.getData("doc/税费管理20141105100144.xls", 0);
			
			if (array != null && array.length > 0){
				for(int i = 0; i < array.length; i++){
					System.out.print(i + ": ");
					for(int j = 0; j < array[i].length; j ++){
						System.out.print(array[i][j] + " ");
					}
					System.out.println();
				}
				
				setData(array, "税费信息");
			}
			
			
		} catch (FileNotFoundException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
	}
	
	/**
	 * 测试写入数据
	 * @param data 数据
	 * @param sheetName sheet名称
	 */
	public static void setData(String [][] data, String sheetName){
		String fileName="doc/税费管理20141105100144Copy.xls";
		try {
			ExcelHelper.setData(fileName, data, sheetName);
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
	}
}
