package generics;

import java.io.FileInputStream;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

public class Excel 
{
	public static String getCellValue(String xpath, String sheet, int row, int cell)
	{
		String v="";
		try
		{
			FileInputStream fis=new FileInputStream(xpath);
			Workbook wb=WorkbookFactory.create(fis);
			v=wb.getSheet(sheet).getRow(row).getCell(cell).toString();
			
			
		}
		catch(Exception e)
		{
			
		}
		return v;
	}
	
	public static int getRowCount(String xpath, String sheet)
	{
		int v=0;
		try
		{
			FileInputStream fis=new FileInputStream(xpath);
			Workbook wb=WorkbookFactory.create(fis);
			v=wb.getSheet(sheet).getLastRowNum();
	    }
		catch(Exception e)
		{
			
		}
		return v;
	}

}
