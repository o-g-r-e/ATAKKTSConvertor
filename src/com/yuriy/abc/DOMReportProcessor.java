package com.yuriy.abc;

import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellRangeAddress;

public class DOMReportProcessor {

	public List<RowEntity> rows;
	
	static class RowDetector
	{
		public static boolean isHead(RowEntity row)
		{
			if(row.getCells()[0].getData().contains("Дата / Время") && 
			   row.getCells()[2].getData().contains("Код") &&
			   row.getCells()[4].getData().contains("Класс события") &&
			   row.getCells()[6].getData().contains("ШП") &&
			   row.getCells()[7].getData().contains("Описание события") &&
			   row.getCells()[10].getData().contains("Канал"))
			{
				return true;
			}
			
			return false;
		}
		
		public static boolean isKTSinRow(RowEntity row)
		{
			//for (int i = 0; i < row.getCells().length; i++) {
				if(row.getCells()[5].getData().contains("проверка КТС"))
					return true;
			//}
			
			return false;
		}
		
		public static boolean isEmptyRow(RowEntity row)
		{
			if(row.getCells().length == 0)
				return true;
			
			for (int i = 0; i < row.getCells().length; i++) {
				if(!row.getCells()[i].getData().equals(""))
					return false;
			}
			
			return true;
		}
		
		public static boolean isPageSeparatorRow(RowEntity row)
		{
			if(row.getCells()[3].getData().contains("Тревоги за период"))
				return true;
			
			return false;
		}
	}
	
	public DOMReportProcessor()
	{
		rows = new ArrayList<RowEntity>();
	}
	
	public void process(Sheet srcSheet)
	{
		System.out.print("Converting file to array... ");
		rows = new ArrayList<RowEntity>(Arrays.asList(sheetRowsToArray(srcSheet)));
		System.out.print("Done.\n");
		
		System.out.print("Detect and remove extra tables... ");
		ExcelRange[] extraTablesIndices =  detectExtraTables(rows.toArray(new RowEntity[0]));
		removeAreasFromList(extraTablesIndices, rows, null);
		System.out.print("Done.\n");
		
		System.out.print("Remove page separators... ");
		removeSepRows(rows);
		System.out.print("Done.\n");
		
		System.out.print("Remove empty rows... ");
		removeEmptyRows(rows);
		System.out.print("Done.\n");
		
		System.out.print("Restore some empty rows... ");
		restoreSomeEmptyRows(rows);
		System.out.print("Done.\n");
		
		System.out.print("Detect and remove no KTS tables... ");
		ExcelRange[] grayTablesIndices =  detectGrayTables(rows.toArray(new RowEntity[0]));
		int[] noKTSRangesIndices = detectNoKTSRanges(grayTablesIndices, rows.toArray(new RowEntity[0]));
		removeAreasFromList(grayTablesIndices, rows, noKTSRangesIndices);
		System.out.print("Done.\n");
		
		System.out.print("Detect and remove empty ATAKs... ");
		ExcelRange[] emptyATAKsIndices =  detectEmptyATAKs(rows.toArray(new RowEntity[0]));
		removeAreasFromList(emptyATAKsIndices, rows, null);
		System.out.print("Done.\n");
		
		System.out.print("Remove extra rows... ");
		removeExtraRows(rows);
		System.out.print("Done.\n");
		
		System.out.print("Modify titles... ");
		modifyATAKsTitles(rows.toArray(new RowEntity[0]));
		modifyTitle(rows.toArray(new RowEntity[0]), 0, 0);
		System.out.print("Done.\n");
		
		System.out.println("Complete.");
	}
	
	private void removeAreasFromList(ExcelRange[] areas, List<RowEntity> list, int[] areasIndices)
	{
		if(areasIndices != null && areasIndices.length > 0)
		{
			for (int i = areasIndices.length-1; i >= 0; i--) {
				list.subList(areas[areasIndices[i]].startIndex, areas[areasIndices[i]].endIndex).clear();
			}
		}
		else
		{
			for (int i = areas.length-1; i >= 0; i--) {
				list.subList(areas[i].startIndex, areas[i].endIndex).clear();
			}
		}
	}
	
	private RowEntity[] sheetRowsToArray(Sheet srcSheet)
	{
		RowEntity[] resultArray = new RowEntity[srcSheet.getLastRowNum()];
		Row currentRow = null;
		for (int i = 0; i < srcSheet.getLastRowNum(); i++) {
			currentRow = srcSheet.getRow(i);
			int rowLength = currentRow.getLastCellNum();
			if(rowLength < 0)
				rowLength = 0;
			CellEntity[] cells = new CellEntity[rowLength];
			for (int j = 0; j < cells.length; j++) {
				cells[j] = new CellEntity(currentRow.getCell(j).getStringCellValue(), currentRow.getCell(j).getCellStyle());
			}
			
			CellRangeAddress[] ranges = getRowMergedRegions(srcSheet, i);
			CellRange[] rowRanges = new CellRange[ranges.length];
			
			for (int j = 0; j < rowRanges.length; j++) {
				rowRanges[j] = new CellRange(ranges[j].getFirstRow(), ranges[j].getLastRow(), ranges[j].getFirstColumn(), ranges[j].getLastColumn());
			}
			
			resultArray[i] = new RowEntity(cells, rowRanges);
		}
		
		return resultArray;
	}
	
	private void modifyTitle(RowEntity[] rows, int titleRowIndex, int titleCellIndex)
	{
		String title = rows[titleRowIndex].getCells()[titleCellIndex].getData();
		title = "Проверки КТС за "+title.substring(18, title.length());
		rows[titleRowIndex].getCells()[titleCellIndex].setData(title);
	}
	
	private void modifyATAKsTitles(RowEntity[] rows)
	{
		String title = null;
		for (int i = 0; i < rows.length; i++) {
			title = rows[i].getCells()[1].getData();
			if(title.contains("М-Н"))
			{
				rows[i].getCells()[1].setData(title.substring(title.indexOf("М-Н"), title.indexOf("М-Н")+15));
			}
		}
	}
	
	private ExcelRange[] detectEmptyATAKs(RowEntity[] rows)
	{
		List<ExcelRange> indices = new ArrayList<ExcelRange>();
		
		for (int i = 0; i < rows.length; i++) {
			if(rows[i].getCells()[1].getData().contains("М-Н") && RowDetector.isEmptyRow(rows[i+4]))
			{
				indices.add(new ExcelRange(i, i+5));
			}
		}
		
		ExcelRange[] result = new ExcelRange[indices.size()];
		for (int i = 0; i < result.length; i++) {
			result[i] = indices.get(i);
		}
		
		return result;
	}
	
	private void restoreSomeEmptyRows(List<RowEntity> rows)
	{
		CellEntity[] emptyRow = new CellEntity[11];
		for (int i = 0; i < emptyRow.length; i++) {
			emptyRow[i] = new CellEntity("", null);
		}
		rows.add(1, new RowEntity(emptyRow, new CellRange[]{}));
		for (int i = rows.size()-1; i > 2; i--) {
			if(rows.get(i).getCells()[1].getData().contains("М-Н"))
			{
				rows.add(i, new RowEntity(emptyRow, new CellRange[]{}));
			}
		}
	}
	
	private void removeSepRows(List<RowEntity> rows)
	{
		for (int i = rows.size()-1; i >= 0; i--) {
			if(!RowDetector.isEmptyRow(rows.get(i)) && rows.get(i).getCells()[3].getData().contains("Тревоги за период"))
			{
				rows.remove(i);
			}
		}
	}
	
	private void removeExtraRows(List<RowEntity> rows)
	{
		for (int i = rows.size()-1; i >= 0; i--) {
			if(!RowDetector.isEmptyRow(rows.get(i)) && (rows.get(i).getCells()[5].getData().contains("Вызов группы") || rows.get(i).getCells()[5].getData().contains("Прибытие группы")))
			{
				rows.remove(i);
			}
		}
	}
	
	private void removeEmptyRows(List<RowEntity> rows)
	{
		for (int i = rows.size()-1; i >= 0; i--) {
			if(RowDetector.isEmptyRow(rows.get(i)))
			{
				rows.remove(i);
			}
		}
	}
	
	private ExcelRange[] detectGrayTables(RowEntity[] rows)
	{
		List<ExcelRange> indices = new ArrayList<ExcelRange>();
		int startIndex = 0;
		boolean foundNewTable = false;
		for (int i = 0; i < rows.length; i++) {
			if(!foundNewTable)
			{
				if(RowDetector.isHead(rows[i]))
				{
					startIndex = i;
					foundNewTable = true;
				}
			}
			else
			{
				if(RowDetector.isHead(rows[i]) || RowDetector.isEmptyRow(rows[i]))
				{
					indices.add(new ExcelRange(startIndex, i));
					if(RowDetector.isHead(rows[i]))
					{
						startIndex = i;
					}
					else
					{
						foundNewTable = false;
					}
				}
			}
		}
		
		ExcelRange[] result = new ExcelRange[indices.size()];
		for (int i = 0; i < result.length; i++) {
			result[i] = indices.get(i);
		}
		
		return result;
	}
	
	private int[] detectNoKTSRanges(ExcelRange[] ranges, RowEntity[] rows)
	{
		List<Integer> removeIndices = new ArrayList<Integer>();
		for (int i = 0; i < ranges.length; i++) {
			boolean ktsFound = false;
			for (int j = ranges[i].startIndex; j < ranges[i].endIndex; j++) {
				if(RowDetector.isKTSinRow(rows[j]))
				{
					ktsFound = true;
				}
			}
			
			if(!ktsFound)
			{
				removeIndices.add(i);
			}
		}
		
		int[] result = new int[removeIndices.size()];
		for (int i = 0; i < result.length; i++) {
			result[i] = removeIndices.get(i);
		}
		
		return result;
	}
	
	private ExcelRange[] detectExtraTables(RowEntity[] rows)
	{
		List<ExcelRange> indices = new ArrayList<ExcelRange>();
		RowEntity row;
		ExcelRange range = null;
		for (int i = 0; i < rows.length; i++) {
			row = rows[i];
			
			if(row.getCells().length <= 0)
				continue;
			
			if(row.getCells()[0].getData().contains("Телефоны"))
			{
				range = new ExcelRange();
				range.startIndex = i + 1;
				indices.add(range);
			}
			
			if( row.getCells()[0].getData().contains("Всего событий"))
			{
				range.endIndex = i + 2;
			}
		}
		
		ExcelRange[] result = new ExcelRange[indices.size()];
		for (int i = 0; i < result.length; i++) {
			result[i] = indices.get(i);
		}
		return result;
	}
	
	public void removeAllRows(Sheet srcSheet)
	{
		while(srcSheet.getLastRowNum() > 0)
		{
			removeRow(srcSheet, srcSheet.getLastRowNum(), false, true);
		}
		removeRow(srcSheet, srcSheet.getLastRowNum(), false, true);
	}
	
	public void pasteRows(Sheet srcSheet, RowEntity[] rows)
	{
		for (int i = 0; i < rows.length; i++) {
			Row newRow = srcSheet.createRow(i);
			if(rows[i].getCells().length <= 0)
				continue;
			
			for (int j = 0; j < rows[i].getAreas().length; j++) {
				rows[i].getAreas()[j].setFirstRow(i);
				rows[i].getAreas()[j].setLastRow(i);
				
				srcSheet.addMergedRegion(new CellRangeAddress(rows[i].getAreas()[j].getFirstRow(), 
															  rows[i].getAreas()[j].getLastRow(), 
															  rows[i].getAreas()[j].getFirstCol(), 
															  rows[i].getAreas()[j].getLastCol()));
			}
			
			for (int j = 0; j < 11; j++) {
				Cell cell = newRow.createCell(j);
				
				if(rows[i].getCells()[j].getStyle() != null)
					cell.setCellStyle(rows[i].getCells()[j].getStyle());
				
				cell.setCellValue(rows[i].getCells()[j].getData());
			}
		}
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
	
	private CellRangeAddress[] getRowMergedRegions(Sheet sheet, int row)
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
}
