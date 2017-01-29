package com.yuriy.abc;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Sheet;

public class Main {
	
	
	 
	public static void main(String[] args) throws FileNotFoundException, IOException {
		/*String srcFileName = null;
		String defaultFileName = "atak.xls";
		
		if(args.length == 0 || args[0].length() == 0)
		{
			File f = new File(defaultFileName);
			if(f.exists() && !f.isDirectory()) { 
				srcFileName = "atak.xls";
			}
		}
		else
		{
			srcFileName = args[0];
		}
		if(srcFileName != null)
		{
			HSSFWorkbook workBook = new HSSFWorkbook(new FileInputStream(srcFileName));
			Sheet srcSheet = workBook.getSheetAt(0);
			DOMReportProcessor proc = new DOMReportProcessor();
			proc.process(srcSheet);
			proc.removeAllRows(srcSheet);
			proc.pasteRows(srcSheet, proc.rows.toArray(new RowEntity[0]));
			
			FileOutputStream out = new FileOutputStream(srcFileName);
			workBook.write(out);
			out.close();
		}*/
		
		new ViewProcessor(new DOMReportProcessor(), new MainWindow());
	}
}