package testRunners;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.nio.file.Paths;
import java.nio.file.StandardOpenOption;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.Iterator;
import java.util.List;
import java.util.Map;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;

import org.apache.commons.collections.bag.SynchronizedSortedBag;
import org.apache.poi.openxml4j.util.ZipSecureFile;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFCreationHelper;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Xls_ReaderAsia {
	//public static Properties prop;
	public String path;
	public FileInputStream fis = null;
	public FileOutputStream fileOut = null;
	private XSSFWorkbook workbook = null;
	private XSSFSheet sheet = null;
	private XSSFRow row = null;
	private XSSFCell cell = null;
	
	public Xls_ReaderAsia(String path) {
		
		this.path = path;
		try {
			fis = new FileInputStream(path);
			ZipSecureFile.setMinInflateRatio(0);
			workbook = new XSSFWorkbook(fis);
			sheet = workbook.getSheetAt(0);
			fis.close();
		} catch (Exception e) {
			// TODO Auto-generated catch block
			//e.printStackTrace();
		}
	}
	// returns the row count in a sheet

	public int getRowCount(String sheetName) {
		int index = workbook.getSheetIndex(sheetName);
		if (index == -1)
			return 0;
		else {
			sheet = workbook.getSheetAt(index);
			int number = sheet.getLastRowNum() + 1;
			return number;
		}

	}
	
	
	// returns true if data is set successfully else false
	public boolean setCellData(String filename, String sheetName, String colName, int rowNum, String data) throws Exception {
		FileInputStream file = null;
	       
        FileOutputStream out = null;
        Cell cell;
        Workbook workbook = null;
    try{
         file = new FileInputStream(new File(filename));

         workbook = WorkbookFactory.create(file);
         Sheet sheet = workbook.getSheet(sheetName);
         Row row = sheet.getRow(0);				


			if (rowNum <= 0)
				return false;

			int index = workbook.getSheetIndex(sheetName);
			int colNum = -1;
			if (index == -1)
				return false;

			sheet = workbook.getSheetAt(index);
		//	CellStyle cellStyle = workbook.createCellStyle();
			//row = sheet.getRow(0);
			for (int i = 0; i < row.getLastCellNum(); i++) {
				// System.out.println(row.getCell(i).getStringCellValue().trim());
				if (row.getCell(i).getStringCellValue().trim().equals(colName)){
					colNum = i;
					break;
				
			}}
			if (colNum == -1)
				return false;

			//sheet.autoSizeColumn(colNum);
			row = sheet.getRow(rowNum - 1);
			if (row == null)
				row = sheet.createRow(rowNum - 1);

			cell = row.getCell(colNum);
			if (cell == null)
				cell = row.createCell(colNum);

			// cell style
			 CellStyle cs = workbook.createCellStyle();
			// cs.setWrapText(true);
			// cell.setCellStyle(cs);			
			cell.setCellValue(data);
		//	CellUtil.setCellStyleProperty(cell, workbook, "CellRow", CellStyle.VERTICAL_CENTER);
			
		} catch (Exception e) {
			e.printStackTrace();
			return false;
		}
		finally{
	         out=new FileOutputStream(new File(filename));
	         workbook.write(out);	         
	         out.close();
	         file.close();
		}
		return true;
	}
	public boolean setCellDataInt(String sheetName, String colName, int rowNum, int data) throws Exception {
		try {
			fis = new FileInputStream(path);
			workbook = new XSSFWorkbook(fis);

			if (rowNum <= 0)
				return false;

			int index = workbook.getSheetIndex(sheetName);
			int colNum = -1;
			if (index == -1)
				return false;

			sheet = workbook.getSheetAt(index);
			CellStyle cellStyle = workbook.createCellStyle();
			row = sheet.getRow(0);
			for (int i = 0; i < row.getLastCellNum(); i++) {
				// System.out.println(row.getCell(i).getStringCellValue().trim());
				if (row.getCell(i).getStringCellValue().trim().equals(colName)){
					colNum = i;
					break;
			}}
			if (colNum == -1)
				return false;

			//sheet.autoSizeColumn(colNum);
			row = sheet.getRow(rowNum - 1);
			if (row == null)
				row = sheet.createRow(rowNum - 1);

			cell = row.getCell(colNum);
			if (cell == null)
				cell = row.createCell(colNum);

			// cell style
			// CellStyle cs = workbook.createCellStyle();
			// cs.setWrapText(true);
			// cell.setCellStyle(cs);			
			cell.setCellValue(data);
			fileOut = new FileOutputStream(path);
			workbook.write(fileOut);
			fileOut.close();
			
		} catch (Exception e) {
			e.printStackTrace();
			return false;
		}
		finally{
			fis.close();
			fileOut.close();

		}
		return true;
		
	}
	public boolean autoSizeColumns(String sheetName) throws Exception {
		fis = new FileInputStream(path);
		workbook = new XSSFWorkbook(fis);

		int rowNum = 0;
		if (rowNum <= 0)
			return false;

		int index = workbook.getSheetIndex(sheetName);
		int colNum = -1;
		if (index == -1)
			return false;

		sheet = workbook.getSheetAt(index);

	        if (((XSSFSheet) sheet).getPhysicalNumberOfRows() > 0) {
	            Row row = ((XSSFSheet) sheet).getRow(((XSSFSheet) sheet).getFirstRowNum());
	            Iterator<Cell> cellIterator = row.cellIterator();
	            while (cellIterator.hasNext()) {
	                Cell cell = cellIterator.next();
	                int columnIndex = cell.getColumnIndex();
	                ((XSSFSheet) sheet).autoSizeColumn(columnIndex);
	            }
	        }
			return false;
	    }
	

	// returns true if sheet is created successfully else false
	public boolean addSheet(String sheetname) {

		FileOutputStream fileOut;
		try {
			workbook.createSheet(sheetname);
			fileOut = new FileOutputStream(path);
			workbook.write(fileOut);
			fileOut.close();
		} catch (Exception e) {
			e.printStackTrace();
			return false;
		}
		return true;
	}

	// returns true if sheet is removed successfully else false if sheet does
	// not exist
	public boolean removeSheet(String sheetName, String FileName) throws Exception {
		fis = new FileInputStream(FileName);
		workbook = new XSSFWorkbook(fis);
			int index = workbook.getSheetIndex(sheetName);
			if (index == -1)
				return false;

			FileOutputStream fileOut = null;
			try {
				workbook.removeSheetAt(index);
				workbook.setActiveSheet(0);
				fileOut = new FileOutputStream(FileName);
				workbook.write(fileOut);
				fileOut.close();
			//	System.out.println("Sheet is deleted");
				} 
			catch (Exception e) 
			{
			//	e.printStackTrace();
			//addSheet("dummy");
				removeSheet1(fis,workbook,fileOut,sheetName,FileName);
			//	fis.close();
			//	workbook.close();
			//	if(fileOut!=null){
			//	fileOut.close();
			//	}File myFile = new File(FileName);
			//	myFile = myFile.getCanonicalFile();
			//	myFile.delete();
				System.out.println("File is deleted");				
//			Path path = FileSystems.getDefault().getPath(FileName);
//			 Files.delete(path);

			//e.printStackTrace();
			return false;
			}
			return true;
		}

	private boolean removeSheet1(FileInputStream fis,XSSFWorkbook workbook,FileOutputStream fileOut, String sheetName,String FileName) throws Exception {
		//	fis = new FileInputStream(FileName);
		//	workbook = new XSSFWorkbook(fis);
		//	int index = workbook.getSheetIndex(sheetName);
		//	if (index == -1)
		//		return false;
		//	FileOutputStream fileOut = null;
			try {
				fis.close();
				fileOut = new FileOutputStream(FileName);
			//	workbook.write(fileOut);
				workbook.close();
				fileOut.flush();
				fileOut.close();				
				File myFile = new File(FileName);
				myFile = myFile.getCanonicalFile();
				myFile.delete();
				} 
			catch (Exception e) 
			{
				e.printStackTrace();
			}	
			return true;
	}
	
	// returns true if column is created successfully
	public boolean addColumn(String sheetName, String colName) {
		// System.out.println("**************addColumn*********************");
		try {
			fis = new FileInputStream(path);
			workbook = new XSSFWorkbook(fis);
			int index = workbook.getSheetIndex(sheetName);
			if (index == -1)
				return false;
			XSSFCellStyle style = workbook.createCellStyle();
			 XSSFFont my_font=workbook.createFont();
             /* set the weight of the font */
           //  my_font.setBoldweight(HSSFFont.BOLDWEIGHT_BOLD);
             style.setVerticalAlignment(VerticalAlignment.TOP);
             style.setWrapText(true);
             /* attach the font to the style created earlier */
             style.setFont(my_font);
			sheet = workbook.getSheetAt(index);
			row = sheet.getRow(0);
			if (row == null)
				row = sheet.createRow(0);
			// cell = row.getCell();
			// if (cell == null)
			// System.out.println(row.getLastCellNum());
			if (row.getLastCellNum() == -1)
				cell = row.createCell(0);
			else
				cell = row.createCell(row.getLastCellNum());
			cell.setCellValue(colName);
			cell.setCellStyle(style);
			fileOut = new FileOutputStream(path);
			workbook.write(fileOut);
			fileOut.close();

		} catch (Exception e) {
			e.printStackTrace();
			return false;
		}

		return true;

	}
	// removes a column and all the contents
	public boolean removeColumn(String sheetName, int colNum) {
		try {
			if (!isSheetExist(sheetName))
				return false;
			fis = new FileInputStream(path);
			workbook = new XSSFWorkbook(fis);
			sheet = workbook.getSheet(sheetName);
			XSSFCellStyle style = workbook.createCellStyle();
		//	style.setFillForegroundColor(HSSFColor.GREY_40_PERCENT.index);
			XSSFCreationHelper createHelper = workbook.getCreationHelper();
		//	style.setFillPattern(HSSFCellStyle.NO_FILL);

			for (int i = 0; i < getRowCount(sheetName); i++) {
				row = sheet.getRow(i);
				if (row != null) {
					cell = row.getCell(colNum);
					if (cell != null) {
						cell.setCellStyle(style);
						row.removeCell(cell);
					}
				}
			}
			fileOut = new FileOutputStream(path);
			workbook.write(fileOut);
			fileOut.close();
		} catch (Exception e) {
			e.printStackTrace();
			return false;
		}
		return true;

	}

	// find whether sheets exists
	public boolean isSheetExist(String sheetName) {
		int index = workbook.getSheetIndex(sheetName);
		if (index == -1) {
			index = workbook.getSheetIndex(sheetName.toUpperCase());
			if (index == -1)
				return false;
			else
				return true;
		} else
			return true;
	}

	// returns number of columns in a sheet
	public int getColumnCount(String sheetName) {
		// check if sheet exists
		if (!isSheetExist(sheetName))
			return -1;

		sheet = workbook.getSheet(sheetName);
		row = sheet.getRow(0);

		if (row == null)
			return -1;

		return row.getLastCellNum();

	}	
	//public static String TESTDATA_SHEET_PATH=prop.getProperty("XmlFilePath");

	
	public Object[][] getTestData(String sheetName,String ExcelFilePath) throws Exception, Exception {
		FileInputStream file = null;
		Workbook book = null;
		 org.apache.poi.ss.usermodel.Sheet sheet1;
		try {
			file = new FileInputStream(ExcelFilePath);
		} catch (FileNotFoundException e) {
			e.printStackTrace();
		}
		try {
			book = WorkbookFactory.create(file);
		} catch (IOException e) {
			e.printStackTrace();
		}
		sheet1 = book.getSheet(sheetName);
		Object[][] data = new Object[sheet1.getLastRowNum()][sheet1.getRow(0).getLastCellNum()];
		for (int i = 0; i < sheet1.getLastRowNum(); i++) {
			for (int k = 0; k < sheet1.getRow(0).getLastCellNum(); k++) {
				if (sheet1.getRow(i + 1).getCell(k)!= null) {
				data[i][k] = sheet1.getRow(i + 1).getCell(k).toString();
			}}
		}
		return data;
	}
	

	/*public static Cell getTestDatamap23(String sheetName,String ExcelFilePath,String Lni,int cellnumber) throws Exception, Exception {
		FileInputStream file = null;
		Workbook book23 = null;
		org.apache.poi.ss.usermodel.Sheet sheet23;
		int getrow = 0;
		 Cell celldata = null;
		try {
			file = new FileInputStream(ExcelFilePath);
			book23 = WorkbookFactory.create(file);
		
		sheet23 = book23.getSheet(sheetName);
		Iterator<Row> iterator = sheet23.iterator();

        while (iterator.hasNext()) {
        	Row currentRow = iterator.next();
            Iterator<Cell> cellIterator = currentRow.iterator();
            while (cellIterator.hasNext()) {

                Cell currentCell = cellIterator.next();
               
				if (currentCell.getCellType() == Cell.CELL_TYPE_STRING) {
                	if(currentCell.getStringCellValue().contains(Lni)){
                   // System.out.print(currentCell.getStringCellValue() + "--");
                     getrow=currentCell.getRowIndex();
                     celldata = sheet23.getRow(getrow).getCell(cellnumber);

//                     System.out.print("first value is-->"+String.valueOf(currentCell.getStringCellValue())+ "--");
                    break;
                }} else if (currentCell.getCellType() == Cell.CELL_TYPE_NUMERIC) {
                	if(String.valueOf(currentCell.getNumericCellValue()).contains(Lni)){
                		 getrow=currentCell.getRowIndex();
                		 celldata = sheet23.getRow(getrow).getCell(cellnumber);
  //                   System.out.print("value is-->"+String.valueOf((int)currentCell.getNumericCellValue())+ "--");
                    break;
                }}

            }
         //   System.out.println();
        }
		}catch(Exception e){}
		finally{
			file.close();
		}
		
		return celldata;
    }
        
*/	
	
	
	
	public int columnName(String Sheet,String ColumnName,String filename) throws Exception {
		int coefficient = 0;	
		Workbook book = null;
		org.apache.poi.ss.usermodel.Sheet sheet1;
	    FileInputStream inputStream = null;
		
		try{
			inputStream	 = new FileInputStream(new File(filename));
	    book = WorkbookFactory.create(inputStream);
	    sheet1 = book.getSheet(Sheet);
	    Row row = sheet1.getRow(0);
	    int cellNum = row.getPhysicalNumberOfCells();
	    for (int i = 0; i < cellNum; i++) {
	        if ((row.getCell(i).toString()).equals(ColumnName)) {
	            coefficient = i;
	        }
	    }}
	    catch(Exception e){}
	    finally{inputStream.close();}
	    return coefficient;
	}	
	public boolean getcolumnName(String Sheet,String ColumnName,String filename) throws Exception {
		Workbook book = null;
		org.apache.poi.ss.usermodel.Sheet sheet1;
	   
		boolean flag = false;	
	    FileInputStream inputStream = new FileInputStream(new File(filename));
	    book = WorkbookFactory.create(inputStream);
	    sheet1 = book.getSheet(Sheet);
	    Row row = sheet1.getRow(0);
	    int cellNum = row.getPhysicalNumberOfCells();
	    for (int i = 0; i < cellNum; i++) {
	        if ((row.getCell(i).toString()).equals(ColumnName)) {
	        	flag=true;
	        }
	    }

	    return flag;
	}
	public List<String> SheetsName(String filename) throws Exception {
		List<String> sheetNames = new ArrayList<String>();
		 FileInputStream fileInputStream = null;
	        try {
	        	FileInputStream inputStream = new FileInputStream(new File(filename));
	    	    Workbook wb = WorkbookFactory.create(inputStream);
	    	            for (int i = 0; i < wb.getNumberOfSheets(); i++) {

	               // System.out.println("Sheet name: " + wb.getSheetName(i));
	                sheetNames.add(wb.getSheetName(i));
	            }

	        } catch (IOException e) {
	            e.printStackTrace();
	        } finally {
	            if (fileInputStream != null) {
	                try {
	                    fileInputStream.close();
	                } catch (IOException e) {
	                    e.printStackTrace();
	                }
	            }
	        }
			return sheetNames;
	    }

	public int addRowinSheet(String sheetName,String filename) throws Exception {
		Workbook book = null;
		org.apache.poi.ss.usermodel.Sheet sheet1;
		int lastRow = 0 ;
	    FileInputStream inputStream = new FileInputStream(new File(filename));
	 try{
	    book = WorkbookFactory.create(inputStream);
	    sheet1 = book.getSheet(sheetName);
	    
		lastRow = sheet1.getLastRowNum(); 
		  
		  //  if (lastRow < startRow) {
		sheet1.createRow(1);
	}	
		catch(Exception e){}
	    finally{
	        if(inputStream!=null)
	        	inputStream.close();
	       
	    }
			return lastRow;
		}
	public void removeRowsFromSheet(String sheetName,String filename) throws Exception {
		FileInputStream file = null;
       
        FileOutputStream out = null;
    try{
         file = new FileInputStream(new File(filename));

         Workbook wb = WorkbookFactory.create(file);
         Sheet sheet = wb.getSheet(sheetName);

        for (int i = 1; i < sheet.getLastRowNum() + 1; i++) {
            Row row = sheet.getRow(i);
            sheet.removeRow(row);
            }
         out = new FileOutputStream(new File(filename));
        wb.write(out);
    }
    catch(Exception e){}
    finally{
        if(file!=null)
        file.close();
        if(out!=null)
        out.close();
        }

    }
	public void removeEmptyRows(String sheetName,String filename) throws Exception {
		FileInputStream file = null;       
        FileOutputStream out = null;
    try{
         file = new FileInputStream(new File(filename));

         Workbook wb = WorkbookFactory.create(file);
         Sheet sheet = wb.getSheet(sheetName);
         Boolean isRowEmpty = Boolean.FALSE;
		    for(int i = 0; i <= sheet.getLastRowNum(); i++){
		      if(sheet.getRow(i)==null){
		        isRowEmpty=true;
		        sheet.shiftRows(i + 1, sheet.getLastRowNum()+1, -1);
		        i--;
		        continue;
		      }
		      for(int j =0; j<sheet.getRow(i).getLastCellNum();j++){
		        if(sheet.getRow(i).getCell(j) == null || 
		        sheet.getRow(i).getCell(j).toString().trim().equals("")){
		          isRowEmpty=true;
		        }else {
		          isRowEmpty=false;
		          break;
		        }
		      }
		      if(isRowEmpty==true){
		        sheet.shiftRows(i + 1, sheet.getLastRowNum()+1, -1);
		        i--;
		      }
		    }
         out = new FileOutputStream(new File(filename));
        wb.write(out);
    }
    catch(Exception e){}
    finally{
        if(file!=null)
        file.close();
        if(out!=null)
        out.close();
        }

    }


	public void removeEmptyRows(String sheetName) throws Exception {
	    try{ 
	    	int index = workbook.getSheetIndex(sheetName);
			sheet = workbook.getSheetAt(index);
			Boolean isRowEmpty = Boolean.FALSE;
		    for(int i = 0; i <= sheet.getLastRowNum(); i++){
		      if(sheet.getRow(i)==null){
		        isRowEmpty=true;
		        sheet.shiftRows(i + 1, sheet.getLastRowNum()+1, -1);
		        i--;
		        continue;
		      }
		      for(int j =0; j<sheet.getRow(i).getLastCellNum();j++){
		        if(sheet.getRow(i).getCell(j) == null || 
		        sheet.getRow(i).getCell(j).toString().trim().equals("")){
		          isRowEmpty=true;
		        }else {
		          isRowEmpty=false;
		          break;
		        }
		      }
		      if(isRowEmpty==true){
		        sheet.shiftRows(i + 1, sheet.getLastRowNum()+1, -1);
		        i--;
		      }
		    }
	    }
	    catch(Exception e){}
	    }
		public void correctRowsinSheet(String sheetName,String filename, String columnName,int columnnumber) throws Exception {

			FileInputStream file = null;
		       
	        FileOutputStream out = null;
	    try{
	         file = new FileInputStream(new File(filename));

	         Workbook wb = WorkbookFactory.create(file);
	         Sheet sheet = wb.getSheet(sheetName);
	    	  int colNum=columnnumber;		    
	         Row row = sheet.getRow(1);				
	         
	 	   Cell cell = row.getCell(columnnumber);
	         
	         
	         
	         for(int i = 1; i <= sheet.getLastRowNum(); i++){
		     // if(sheet.getRow(i)!=null){
	        	 System.out.println("value of i" + i);
		    	  cell.setCellValue(i);
		      //}
		    }
	    
	      out = new FileOutputStream(new File(filename));
	        wb.write(out);
	    }
	    catch(Exception e){}
	    finally{
	        if(file!=null)
	        file.close();
	        if(out!=null)
	        out.close();
	        }

		}
		public int getRowCountWithFile(String sheetName,String filename) throws Exception {
			FileInputStream file = null;
			int totleRow = 0;
			try {
				file = new FileInputStream(filename);
			} catch (FileNotFoundException e) {
				e.printStackTrace();
			}
			try {
				workbook = (XSSFWorkbook) WorkbookFactory.create(file);
			
				sheet = workbook.getSheetAt(0);

				totleRow=sheet.getPhysicalNumberOfRows();
			
				return totleRow;
			} catch (IOException e) {
				e.printStackTrace();
			}
			finally{
				file.close();
			}
			return totleRow;
	}
		
		public int addRowinSheet(String sheetName) {
			int index = workbook.getSheetIndex(sheetName);
			int startRow = 1;
		  	sheet = workbook.getSheetAt(index);
			    int lastRow = sheet.getLastRowNum();
			  //  if (lastRow < startRow) {
			        sheet.createRow(startRow);
			    
				return lastRow;
			}
		public int addRowinSheetbySize(String sheetName,int size) {
			int index = workbook.getSheetIndex(sheetName);
			int startRow = 1;
		  	sheet = workbook.getSheetAt(index);
			    int lastRow = sheet.getLastRowNum();
			    if (lastRow < size) {
			        sheet.createRow(size);
			}
			    System.out.println("Row are created");
				return lastRow;
		}
		public static Cell getTestDatamap23(String sheetName,String ExcelFilePath,String Lni,int cellnumber) throws Exception, Exception {
			FileInputStream file = null;
			Workbook book23 = null;
			org.apache.poi.ss.usermodel.Sheet sheet23;
			int getrow = 0;
			 Cell celldata = null;
			try {
				file = new FileInputStream(ExcelFilePath);
				book23 = WorkbookFactory.create(file);
			
			sheet23 = book23.getSheet(sheetName);
			Iterator<Row> iterator = sheet23.iterator();

	        while (iterator.hasNext()) {
	        	Row currentRow = iterator.next();
	            Iterator<Cell> cellIterator = currentRow.iterator();
	            while (cellIterator.hasNext()) {

	                Cell currentCell = cellIterator.next();
	                if (currentCell.getCellType() == CellType.STRING) {
						if(currentCell.getStringCellValue().contains(Lni)){
	                   // System.out.print(currentCell.getStringCellValue() + "--");
	                     getrow=currentCell.getRowIndex();
	                     celldata = sheet23.getRow(getrow).getCell(cellnumber);

//	                     System.out.print("first value is-->"+String.valueOf(currentCell.getStringCellValue())+ "--");
	                break;
	                }} else if (currentCell.getCellType() == CellType.NUMERIC) {
	                	if(String.valueOf(currentCell.getNumericCellValue()).contains(Lni)){
	                		 getrow=currentCell.getRowIndex();
	                		 celldata = sheet23.getRow(getrow).getCell(cellnumber);
	  //                   System.out.print("value is-->"+String.valueOf((int)currentCell.getNumericCellValue())+ "--");
	                    break;
	                }}

	            }
	         //   System.out.println();
	        }
			}catch(Exception e){}
			finally{
				file.close();
			}
			
			return celldata;
	    }
		public static Object[][] getTestDatamap24(String sheetName,String ExcelFilePath) throws Exception, Exception {
			 Workbook book24 = null;
			 org.apache.poi.ss.usermodel.Sheet sheet24;
			FileInputStream file = null;
			int getrow = 0;
			 Cell celldata = null;
			 Object[][] data = null;
			try {
				file = new FileInputStream(ExcelFilePath);
			} catch (FileNotFoundException e) {
				//e.printStackTrace();
			}
			try {
				book24 = WorkbookFactory.create(file);
				   sheet24 = book24.getSheet(sheetName);
				

			
			sheet24 = book24.getSheet(sheetName);
				data = new Object[sheet24.getLastRowNum()][sheet24.getRow(0).getLastCellNum()];
				for (int i = 0; i < sheet24.getLastRowNum(); i++) {
					for (int k = 0; k < sheet24.getRow(0).getLastCellNum(); k++) {
						if (sheet24.getRow(i + 1).getCell(k)!= null) {
						data[i][k] = sheet24.getRow(i + 1).getCell(k).toString();
					}}
				}
			}
		catch(Exception e){}
			finally{
				file.close();
				book24.close();
			}
			
			return data;	
			
	    }

		public int getFileSize(String fileName) {
			 Path path = Paths.get(fileName);
			 int fileSize=0;
			 try {
		            long bytes = Files.size(path);
		          //  System.out.println(String.format("%,d bytes", bytes));
		           // System.out.println(String.format("%,d kilobytes", bytes / 1024));
		            if(Integer.parseInt(String.format("%,d", bytes/1024))>0){
		            fileSize=1;
		            }
		            else{
		//            	fileSize=0;
		            }
		        } catch (IOException e) {
		  //          e.printStackTrace();
		            fileSize=0;
		        }
			return fileSize;
		}

		public void DelteZeroSizeFile(String fileName) {
			File file = new File(fileName);
			if (file.exists()) {
				
			    file.delete();
			} else {
			    System.err.println(
			        "I cannot find '" + file + "' ('" + file.getAbsolutePath() + "')");
			}
			
		}	
		
		
		


public static Map<Integer, String> getTestDatasummary(String ExcelFilePath,String sheetName,int columnno) throws Exception, Exception {
	Map<Integer, String> map = new HashMap<Integer, String>();
	List<String>data1=new ArrayList<String>();
	File myFile = new File(ExcelFilePath); 
    FileInputStream fis = null;
    try {
        fis = new FileInputStream(myFile);
    } catch (FileNotFoundException e) {
        // TODO Auto-generated catch block
        e.printStackTrace();
    } 
    XSSFWorkbook myWorkBook = null;
    try {
        myWorkBook = new XSSFWorkbook (fis);
    } catch (IOException e) {
        e.printStackTrace();
    } 

    XSSFSheet  sheet1 = myWorkBook.getSheet(sheetName);
	Object[][] data = new Object[sheet1.getLastRowNum()][sheet1.getRow(0).getLastCellNum()];
	for (int i = 0; i < sheet1.getLastRowNum(); i++) {
		//data[i][columnno] = sheet1.getRow(i + 1).getCell(columnno).toString();
		//	map.put(i+1,data[i][columnno].toString());			
	map.put(i+1,sheet1.getRow(i + 1).getCell(columnno).toString());
	}
	//System.out.println(data1);
	
	return map;
}

public static int getRowsCount(String excelPath,String sheetName) throws IOException {
    XSSFWorkbook workbook = null;
    XSSFSheet sheet = null;
    int number=0;
    try {
        FileInputStream file = new FileInputStream(new File(excelPath));
        workbook = new XSSFWorkbook(file);
        sheet = workbook.getSheet(sheetName);
        
        int index = workbook.getSheetIndex(sheetName);
		if (index == -1)
			return 0;
		else {
			sheet = workbook.getSheetAt(index);
			number = sheet.getLastRowNum() + 1;
			return number;
		}


    } catch(Exception e) {
        throw e;
    } finally {
        if(workbook != null)
            workbook.close();
    }
}


public boolean deleteRow(String excelPath,String sheetName, int rowNo) throws IOException {
    XSSFWorkbook workbook = null;
    XSSFSheet sheet = null;
    try {
        FileInputStream file = new FileInputStream(new File(excelPath));
        workbook = new XSSFWorkbook(file);
        sheet = workbook.getSheet(sheetName);
        if (sheet == null) {
            return false;
        }
        int lastRowNum = sheet.getLastRowNum();
        if (rowNo >= 0 && rowNo < lastRowNum) {
            sheet.shiftRows(rowNo + 1, lastRowNum, -1);
        }
        if (rowNo == lastRowNum) {
            XSSFRow removingRow=sheet.getRow(rowNo);
            if(removingRow != null) {
                sheet.removeRow(removingRow);
            }
        }
        file.close();
        FileOutputStream outFile = new FileOutputStream(new File(excelPath));
        workbook.write(outFile);
        outFile.close();


    } catch(Exception e) {
        throw e;
    } finally {
        if(workbook != null)
            workbook.close();
    }
    return false;
}



}		