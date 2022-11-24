package com.learning.Task;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileWriter;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.json.simple.JSONArray;
import org.json.simple.JSONObject;

//import ReadXLsheet.Employee;

public class ReadXLSheet {
	@SuppressWarnings("unchecked")
	public static void main(String[] args) throws IOException {
		// TODO Auto-generated method stub

		// Read XLSheet
		String str[][] = new String[11][6];
		int i = 0, j = 0;
		File file = new File(System.getProperty("user.dir")+"/src/test/java/com/learning/repos/EmployeeDetails.xlsx");
		FileInputStream fis = new FileInputStream(file);
		System.out.println("FILE READED");
		XSSFWorkbook book = new XSSFWorkbook(fis);
		XSSFSheet sheet = book.getSheetAt(0);
		XSSFRow row = null;
		XSSFCell cell = null;

		for (i = 0; i < 11; i++) {
			row = sheet.getRow(i);
			for (j = 0; j < 6; j++) {
				cell = row.getCell(j);
				if (String.valueOf(cell.getCellType()) == "STRING") {
					str[i][j] = cell.getStringCellValue();
				} else {
					str[i][j] = String.valueOf(cell.getNumericCellValue());
				}
			}
		}
		for (i = 0; i < 11; i++)
			for (j = 0; j < 6; j++)
				System.out.println(str[i][j]);

		// Convert into List of Map
		System.err.println("\nconvert in list of map\n");
		List<Map<String, String>> listMap = new ArrayList<Map<String, String>>();
		Map<String, String> map = new HashMap<String, String>();
		for (i = 1; i < str.length; i++) {
			map.put("Name", str[i][0]);
			map.put("EmpId", str[i][1]);
			map.put("Designation", str[i][2]);
			map.put("No.of Working Days", str[i][3]);
			map.put("City", str[i][4]);
			map.put("Salary", str[i][5]);
			listMap.add(map);
		}
		for (Map<String, String> map1 : listMap) {
			for (Map.Entry<String, String> a : map1.entrySet()) {
				System.out.println(a.getKey() + "   :" + a.getValue());
			}
			System.out.println("\n");
		}

		// convert into list of class
		List<Employee> list = new ArrayList();

		for (i = 0; i < str.length; i++) {
			Employee emp = new Employee();
			emp.setName(str[i][0]);
			emp.setEmpID(str[i][1]);
			emp.setDesignation(str[i][2]);
			emp.setNOWDays(str[i][3]);
			emp.setCity(str[i][4]);
			emp.setSalary(str[i][5]);
			list.add(emp);
		}

		for (Employee emp1 : list) {
			System.out.print(emp1.getName() + "  ");
			System.out.print(emp1.getEmpID() + "  ");
			System.out.print(emp1.getDesignation() + "  ");
			System.out.print(emp1.getNOWDays() + "  ");
			System.out.print(emp1.getCity() + "  ");
			System.out.print(emp1.getSalary() + "  ");
			System.out.print("\n\n");

		}

		// convert into JSON
		JSONArray jsonArray = new JSONArray();
		for (i = 1; i < str.length; i++) {
			JSONObject jsonObj = new JSONObject();
			jsonObj.put("Name", str[i][0]);
			jsonObj.put("EmpId", str[i][1]);
			jsonObj.put("Designation", str[i][2]);
			jsonObj.put("NOWdays", str[i][3]);
			jsonObj.put("City", str[i][4]);
			jsonObj.put("Salary", str[i][5]);
			jsonArray.add(jsonObj);
		}

		FileWriter file1 = new FileWriter("C:\\Users\\2149352\\eclipse-workspace\\POJO\\outputFile.json");
		file1.write(jsonArray.toJSONString());
		file1.flush();
		file1.close();

	}

}


