package zjr.amy.excel.test;

import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Calendar;
import java.util.List;

import zjr.amy.excel.ExcelHelper;
import zjr.amy.excel.ExcelUtil;
import zjr.amy.model.Person;

public class TestExcelHelper {

	/**
	 * @param args
	 */
	public static void main(String[] args) {
		// TODO Auto-generated method stub
		TestDataModel();
	}

	/**
	 * ���Ի�ȡ����
	 */
	public static void TestGetData() {
		try {
			String[][] array = ExcelHelper.getData(
					"doc/˰�ѹ���20141105100144.xls", 0);

			if (array != null && array.length > 0) {
				for (int i = 0; i < array.length; i++) {
					System.out.print(i + ": ");
					for (int j = 0; j < array[i].length; j++) {
						System.out.print(array[i][j] + " ");
					}
					System.out.println();
				}

				setData(array, "˰����Ϣ");
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
	 * ����д������
	 * 
	 * @param data
	 *            ����
	 * @param sheetName
	 *            sheet����
	 */
	public static void setData(String[][] data, String sheetName) {
		String fileName = "doc/˰�ѹ���20141105100144Copy.xls";
		try {
			ExcelHelper.setData(fileName, data, sheetName);
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
	}

	public static void TestDataModel() {
		Person p = new Person();
		p.setKeyID("1212121212");
		p.setSex("��");
		Calendar date = Calendar.getInstance();
		date.set(2000, 1, 15);
		p.setBirth(date);
		p.setName("����");

		List<Person> list = new ArrayList<Person>();
		list.add(p);
		p = new Person();
		p.setKeyID("1212121212");
		p.setSex("Ů");
		date = Calendar.getInstance();
		date.set(2000, 12, 25);
		p.setBirth(date);
		p.setName("����");
		list.add(p);
		ExcelUtil<Person> bll = new ExcelUtil<Person>();
		String[] headers = { "����", "�Ա�", "����", "��������" };
		try {
			bll.writeData(headers, list, "doc/Person.xls");
		} catch (FileNotFoundException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
	}
}
