import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.util.NumberToTextConverter;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class baseclass_Integers {
	
	public ArrayList<String> getData(String testCaseName) throws IOException
	{
		ArrayList<String> arr = new ArrayList<String>();
		
		FileInputStream fis = new FileInputStream("C:\\Users\\c-deepak.jindal\\eclipse-workspace\\Selenium_DataDriven\\src\\DemoData.xlsx"); 		//create a channel which will access the excel file. 
		XSSFWorkbook obj = new XSSFWorkbook(fis);
		int number_of_sheets = obj.getNumberOfSheets(); // Get total number of sheets in the excel
		for(int i = 0; i< number_of_sheets ; i++ ) // Running loop for all sheets
		{
			if(obj.getSheetName(i).equalsIgnoreCase("testdata"))  //Check if we are on the desired sheet name
			{
				XSSFSheet sheet = obj.getSheetAt(i); // checking all sheets one by one. 
				Iterator<Row> rows = sheet.iterator(); // Using iterator method to scan through all the rows
				Row row = rows.next(); // Control is now on the first row of the sheet
				Iterator<Cell> ce = row.cellIterator(); // Scan through all the cells of the identified row. 
				int k=0; // We are using this variable which will store the column number
				int column = 0 ;
				while(ce.hasNext()) 
				{
					Cell value = ce.next(); //Doing this takes control to the first cell of the idemtified row and value of cell is stored
					if(value.getStringCellValue().equalsIgnoreCase("Testcases")) // Checking if the first columnName is the desired or not. 
					{
						column = k;			
					}
					k++;		
				}
				while(rows.hasNext())
				{
					Row r = rows.next();
					if(r.getCell(column).getStringCellValue().equalsIgnoreCase(testCaseName))
					{
						Iterator<Cell> re = r.cellIterator();
						while(re.hasNext())
						{
							Cell c = re.next();
							if(c.getCellTypeEnum()==CellType.STRING)
							{
								arr.add(c.getStringCellValue());
							}
							else
							{
								arr.add(NumberToTextConverter.toText(c.getNumericCellValue()));
							}
						}
					}
				}
			}
			
		}
		return arr;		
	}
	
	public static void main(String[] args) throws IOException
	{
		baseclass_Integers b = new baseclass_Integers(); //Creating the object of the class which has function for pulling out excel data
		ArrayList data = b.getData("Test"); // call the method, passing the test case name as the argument, the test case we wish to execute
		System.out.println(data.get(0)); // Now that we have all the data available in the arraylist, we can grab each cell value and use it in test cases
		System.out.println(data.get(1));
		System.out.println(data.get(2));
		System.out.println(data.get(3));

		
	}

}
