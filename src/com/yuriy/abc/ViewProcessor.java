package com.yuriy.abc;

import java.awt.Desktop;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.PrintWriter;
import java.io.StringWriter;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Sheet;

public class ViewProcessor implements MainWindow.WindowEventsListener {
	
	private DOMReportProcessor reportProcessor;
	private MainWindow mainWindow;
	private File sourceFile;
	
	ViewProcessor(DOMReportProcessor reportProcessor, MainWindow mainWindow)
	{
		this.reportProcessor = reportProcessor;
		this.mainWindow = mainWindow;
		this.mainWindow.addListener(this);
		this.reportProcessor.addObserver(this.mainWindow);
	}

	@Override
	public void onStartButtonClick() {
		new Thread(new Runnable() {
			public void run() {
			   try {
				   HSSFWorkbook workBook = new HSSFWorkbook(new FileInputStream(sourceFile));
				   Sheet srcSheet = workBook.getSheetAt(0);
					
				   reportProcessor.process(srcSheet);
				   reportProcessor.removeAllRows(srcSheet);
				   reportProcessor.pasteRows(srcSheet, reportProcessor.rows.toArray(new RowEntity[0]));
					
				   FileOutputStream out = new FileOutputStream(sourceFile);
				   workBook.write(out);
				   out.close();
					
				   Desktop.getDesktop().open(sourceFile.getParentFile());
					
				   System.exit(0);
			   } catch (Exception e) {
				   //e.printStackTrace();
				   mainWindow.showException(getStackTrace(e));
			   } finally
			   {
				   System.exit(0);
			   }
			}
		}).start();
	}

	@Override
	public void onFlieLoaded(String fileName) {
		sourceFile = new File(fileName);
	}
	
	private String getStackTrace(Throwable throwable) {
	     StringWriter sw = new StringWriter();
	     PrintWriter pw = new PrintWriter(sw, true);
	     throwable.printStackTrace(pw);
	     return sw.getBuffer().toString();
	}
}