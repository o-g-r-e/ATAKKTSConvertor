package com.yuriy.abc;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellRangeAddress;

public class ReportProcessor
{
	private String sourceFileName;
	private HSSFWorkbook workBook;
	//private int lastRowNumber;
	private Sheet srcSheet;
	public int shiftCounter = 0;
	
	public ReportProcessor(String sourceFileName) throws FileNotFoundException, IOException {
		this.sourceFileName = sourceFileName;
		this.workBook = new HSSFWorkbook(new FileInputStream(this.sourceFileName));
		this.srcSheet = this.workBook.getSheetAt(0);
		//this.lastRowNumber = srcSheet.getLastRowNum();
	}
	
	private static void addAreasToRow(Sheet sheet, int rowIndex, CellRangeAddress[] cellRangeAddreses)
	{
		CellRangeAddress cellRangeAddress = null;
		Row row = sheet.getRow(rowIndex);
		for (int i = 0; i < cellRangeAddreses.length; i++) {
			cellRangeAddress = new CellRangeAddress(row.getRowNum(), (row.getRowNum() + (cellRangeAddreses[i].getLastRow() - cellRangeAddreses[i].getFirstRow())), cellRangeAddreses[i].getFirstColumn(), cellRangeAddreses[i].getLastColumn());
			sheet.addMergedRegion(cellRangeAddress);
		}
	}
	
	private static void cellCopy(Cell srcCell, Cell destCell)
	{
        destCell.setCellStyle(srcCell.getCellStyle());
        
        if (srcCell.getCellComment() != null) {
        	destCell.setCellComment(srcCell.getCellComment());
        }
        
        if (srcCell.getHyperlink() != null) {
        	destCell.setHyperlink(srcCell.getHyperlink());
        }
        
        destCell.setCellType(srcCell.getCellType());

        // Set the cell data value
        switch (srcCell.getCellType()) {
            case Cell.CELL_TYPE_BLANK:
            	destCell.setCellValue(srcCell.getStringCellValue());
                break;
            case Cell.CELL_TYPE_BOOLEAN:
            	destCell.setCellValue(srcCell.getBooleanCellValue());
                break;
            case Cell.CELL_TYPE_ERROR:
            	destCell.setCellErrorValue(srcCell.getErrorCellValue());
                break;
            case Cell.CELL_TYPE_FORMULA:
            	destCell.setCellFormula(srcCell.getCellFormula());
                break;
            case Cell.CELL_TYPE_NUMERIC:
            	destCell.setCellValue(srcCell.getNumericCellValue());
                break;
            case Cell.CELL_TYPE_STRING:
            	destCell.setCellValue(srcCell.getRichStringCellValue());
                break;
        }
	}
	
	private static CellRangeAddress[] getRowMergedRegions(Sheet sheet, int row)
	{
		List<CellRangeAddress> cellRangeAddresses = new ArrayList<CellRangeAddress>();
		CellRangeAddress currentCellRangeAddress;
		 for (int i = 0; i < sheet.getNumMergedRegions(); i++)
		 {
			 currentCellRangeAddress = sheet.getMergedRegion(i);
			 if (currentCellRangeAddress.getFirstRow() == row)
			 {
				 cellRangeAddresses.add(currentCellRangeAddress);
			 }
		 }
		 
		 return cellRangeAddresses.toArray(new CellRangeAddress[0]);
	}
	
	private static Integer[] getRowMergedIndices(Sheet sheet, int row)
	{
		List<Integer> cellRangeAddresses = new ArrayList<Integer>();
		CellRangeAddress currentCellRangeAddress;
		 for (int i = 0; i < sheet.getNumMergedRegions(); i++)
		 {
			 currentCellRangeAddress = sheet.getMergedRegion(i);
			 if (currentCellRangeAddress.getFirstRow() == row)
			 {
				 cellRangeAddresses.add(i);
			 }
		 }
		 
		 return cellRangeAddresses.toArray(new Integer[0]);
	}
	
	private static void removeRow(Sheet sheet, int rowIndex, boolean init, boolean removeAreas)
	{
		if(removeAreas)
		{
			Integer[] mergedAreas = getRowMergedIndices(sheet, rowIndex);
			
			for (int i = mergedAreas.length-1; i >= 0; i--) {
				sheet.removeMergedRegion(mergedAreas[i]);
			}
		}
		Row row = sheet.getRow(rowIndex);
		if(row != null)
			sheet.removeRow(row);
		int cellsNum = 11;//row.getLastCellNum();
		
		if(init)
		{
			Row newRow = sheet.createRow(rowIndex);
			
			for (int i = 0; i < cellsNum; i++) {
				newRow.createCell(i);
			}
		}
	}
	
	private static void rowCopy(Row srcRow, Row destRow)
	{
		for (int i = 0; i < 12; i++)
		{
			if(destRow.getCell(i) == null)
			{
				destRow.createCell(i);
			}
		}
			
		/*for (int i = 0; i < 12; i++)
		{
			if(srcRow.getCell(i) == null)
			{
				srcRow.createCell(i);
			}
		}*/
		
		
		for (int i = 0; i < srcRow.getLastCellNum(); i++) {
			cellCopy(srcRow.getCell(i), destRow.getCell(i));
		}
	}
	
	private void shiftUpRows(Sheet sheet, int startRowIndex, int shiftNum)
	{
		for (int i = startRowIndex; i < sheet.getLastRowNum(); i++) {
			//System.out.println(i);
			removeRow(sheet, i-shiftNum, true, true);
			addAreasToRow(sheet, i-shiftNum, getRowMergedRegions(sheet, i));
			if(sheet.getRow(i) == null)
			{
				Row row = sheet.createRow(i);
				for (int j = 0; j < 11; j++) {
					row.createCell(j);
				}
			}
			rowCopy(sheet.getRow(i), sheet.getRow(i-shiftNum));
			removeRow(sheet, i, false, true);
		}
	}
	
	private static void reportHeadHandler(Row currentRow)
	{
		String headValue = currentRow.getCell(0).getStringCellValue();
		String headString = "КТС отчет";
		headValue = headValue.replaceAll("Тревоги за период", headString);
		currentRow.getCell(0).setCellValue(headValue);
	}
	
	private String writeProgress(String prefix, int n, int max)
	{
		int p = Math.round(((float)n*100.0f)/(float)max);
		
		return "\r"+prefix+"..."+p+"%";
	}
	
	public void removeNoKTSObjects()
	{
		Row currentRow = null;
		for (int i = 0; i < srcSheet.getLastRowNum(); i++) {
			
		}
	}
	
	public void doFirstPhase()
	{
		/*for (int i = srcSheet.getLastRowNum(); i >= 1845; i--)
		{
			removeRow(srcSheet, i, false, true);
			System.out.println("Row "+i);
		}
		
		return;*/
		
		Row currentRow = null;
		boolean findingGeneralKTSTable = false;
		for (int i = 0; i < srcSheet.getLastRowNum(); i++) {
			currentRow = srcSheet.getRow(i);
			
			if(currentRow == null || currentRow.getCell(0) == null)
				continue;
			
			if(i==0)
			{
				reportHeadHandler(currentRow);
			}
			
			if(currentRow.getCell(0).getStringCellValue().startsWith("Объект"))
			{
				Cell c = currentRow.getCell(1);
				String s = c.getStringCellValue();
				c.setCellValue(s.substring(s.indexOf("М-Н"), s.indexOf("М-Н")+15));
				//copyRow(myExcelBook, srcSheet, myExcelBook.getSheetAt(1), i, i);
				findingGeneralKTSTable = true;
				continue;
			}
			
			if(findingGeneralKTSTable)
			{
				if(currentRow.getCell(0).getStringCellValue().startsWith("Дата / Время"))
				{
					int startGeneralKTSTableIndex = i;
					
					for (int j=startGeneralKTSTableIndex;  j < srcSheet.getLastRowNum(); j++) {
						currentRow = srcSheet.getRow(j);
						if(currentRow == null || currentRow.getCell(0) == null)
							continue;
						
						if(currentRow.getCell(0).getStringCellValue().startsWith("Всего событий : "))
						{
							int totalEventsTableLength = (j+2)-startGeneralKTSTableIndex;
							
							shiftUpRows(srcSheet, j+2, totalEventsTableLength);
							findingGeneralKTSTable = false;
							break;
						}
					}
				}
			}
			System.out.print(writeProgress("Removing tables", i, srcSheet.getLastRowNum()));
		}
		System.out.println();
	}
	
	public void doSecondPhase()
	{
		int startKTSDetailsIndex = 0;
		boolean findingKTSEventTable = false;
		Row currentRow = null;
		for (int i = 0; i < srcSheet.getLastRowNum(); i++) {
			currentRow = srcSheet.getRow(i);
			if(currentRow == null || currentRow.getCell(0) == null)
				continue;
			
			if(currentRow.getCell(0).getStringCellValue().startsWith("Дата / Время") && !findingKTSEventTable)
			{
				startKTSDetailsIndex = i;
				findingKTSEventTable = true;
				continue;
			}
			
			if(findingKTSEventTable)
			{
				if(currentRow.getCell(0).getStringCellValue().startsWith("Дата / Время"))
				{
					for (int j = i+1; j < srcSheet.getLastRowNum(); j++) {
						currentRow = srcSheet.getRow(j);
						if(currentRow.getCell(5).getStringCellValue().contains("КТС"))
						{
							findingKTSEventTable = false;
							startKTSDetailsIndex = 0;
							break;
						}
						
						if(currentRow.getCell(0).getStringCellValue().startsWith("Дата / Время"))
						{
							shiftUpRows(srcSheet, j, j-startKTSDetailsIndex);
							findingKTSEventTable = false;
							i = 0;//startKTSDetailsIndex;
							break;
						}
					}
				}
			}
			System.out.print(writeProgress("Removing no KTS tables", i, srcSheet.getLastRowNum()));
		}
		System.out.println();
	}
	
	public void doThirdPhase()
	{
		Row currentRow = null;
		for (int i = 0; i < srcSheet.getLastRowNum(); i++) {
			currentRow = srcSheet.getRow(i);
			if(currentRow == null || currentRow.getCell(0) == null)
				continue;
			
			if(currentRow.getCell(3).getStringCellValue().contains("Тревоги за период"))
			{
				int startShiftIndex = i+2;
				int shift = 3;
				if(!isRowEmpty(srcSheet.getRow(i-1)))
				{
					startShiftIndex++;
				}
				
				shiftUpRows(srcSheet, startShiftIndex, shift);
				
				i-=2;
			}
			System.out.print(writeProgress("Removing page separators", i, srcSheet.getLastRowNum()));
		}
		System.out.println();
	}
	
	public void removingLastRows()
	{
		System.out.println("Removing last rows...");
		while (isRowEmpty(srcSheet.getRow(srcSheet.getLastRowNum())))
		{
			removeRow(srcSheet, srcSheet.getLastRowNum(), false, true);
		}
		System.out.println("Done.");
	}
	
	private boolean isRowEmpty(Row row)
	{
		if(row == null)
			return true;
		
		for (int i = 0; i < 11; i++) {
			if(row.getCell(i) == null)
					continue;
			if(!row.getCell(i).getStringCellValue().equals(""))
				return false;
		}
		
		return true;
	}
	
	public void writeChanges() throws IOException
	{
		FileOutputStream out = new FileOutputStream(sourceFileName);
		workBook.write(out);
		out.close();
	}
}
