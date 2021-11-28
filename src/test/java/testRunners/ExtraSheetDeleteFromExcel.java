package testRunners;

import java.io.File;
import java.io.FileInputStream;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Collections;
import java.util.HashSet;
import java.util.List;
import java.util.Properties;
import java.util.function.Predicate;
import java.util.stream.Collectors;

import org.apache.commons.io.FileUtils;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.google.common.collect.Lists;
import com.google.common.collect.Sets;

public class ExtraSheetDeleteFromExcel {
	 public static void main(String[] args) throws Exception {	         	        
				long startTime = System.currentTimeMillis();
		 		InputStream input = new FileInputStream("./src/test/resources/propertyFiles/config.properties");
				Properties prop = new Properties(); 
				prop.load(input);
				String SummaryResultUpdated=prop.getProperty("MasterGDSExcel");
				String summerysheetName=prop.getProperty("MasterGdsSheet");
				String DpsiDataMappingExcel=prop.getProperty("DpsiDataMappingExcel");
				String dipsiMapping=prop.getProperty("Sheet1");
				String batchMapping=prop.getProperty("batchMapping");
				String LoadDate=prop.getProperty("LoadDate");
				String WaveData=prop.getProperty("WaveData");
				String TestDate=prop.getProperty("TestDate");
				Sheet deleteSheets=GetSheetFromFile(dipsiMapping,DpsiDataMappingExcel);
				Sheet otherdatausingDpsisheet=GetSheetFromFile(batchMapping,DpsiDataMappingExcel);
				
			//	Xls_ReaderAsia DpsiMapping = new Xls_ReaderAsia(DpsiDataMappingExcel);
				
		 		File folder = new File("./Results");
				//String SummaryResultUpdated="./Results/SummaryResults/ExcelResults.xlsx";
				//String summerysheetName="Sheet1";		
				Xls_ReaderAsia testData1 = new Xls_ReaderAsia(DpsiDataMappingExcel);	    		
				File[] listOfFiles = folder.listFiles();
	        	List<String> excelFiles = new ArrayList<String>();
	        	List<String> sheetNames = new ArrayList<String>();
	        	for (File file : listOfFiles) {
	        	    if (file.isFile()) {
	        	    if(file.getName().contains(".xlsx")){
	        	    	excelFiles.add(file.getName());
	        	    }
	        	    }
	        	}
	    		for (int i=0; i<excelFiles.size(); i++) {
	        		long startTime1 = System.currentTimeMillis();
		        		sheetNames=testData1.SheetsName("./Results/"+excelFiles.get(i));     		
			        	for (int j=0; j<sheetNames.size(); j++) {
			        		String SheetMatched=DpsiOtherMatchData(deleteSheets,sheetNames.get(j),1);			        				       
			        			Xls_ReaderAsia testDatanew = new Xls_ReaderAsia("./Results/"+excelFiles.get(i));
			        			if(!excelFiles.get(i).trim().equalsIgnoreCase(SheetMatched)){
			        				testDatanew.removeSheet(sheetNames.get(j),"./Results/"+excelFiles.get(i));
			        			//	System.out.println("Sheet is deleted as -->"+sheetNames.get(j));
			        			//	System.out.println("Excel file name is -->"+excelFiles.get(i));
			        			}		        		
			        	long endTime1 = System.currentTimeMillis();
			        	System.out.println("That took " + (endTime1 - startTime1)/1000 + "seconds in Excel--> "+excelFiles.get(i));    
        		}        		
	        	
        	         	   	
	  }
	    		long endTime = System.currentTimeMillis();
	        	System.out.println("That took " + (endTime - startTime)/1000 + "seconds");
	    		System.out.println("Data is updated");  
	 }

	@SuppressWarnings("resource")
	public static Sheet GetSheetFromFile(String excelSheet,String ExcelFile) throws Exception{
		FileInputStream fis = new FileInputStream(ExcelFile);
		XSSFWorkbook  workbook = new XSSFWorkbook(fis);
	          XSSFSheet sheet = workbook.getSheetAt(0);
	        sheet = workbook.getSheet(excelSheet);	    
	        return sheet;
	}
	
	private static String DpsiOtherMatchData(Sheet sheet,String dpsi,int data) throws Exception {
		Object[][] data1=FindDpsiFromSheet(sheet);
		//System.out.println(data1.length);
		
		int i1 =0;
		 for(int p1=0;p1<data1.length;p1++)
	        {
	     if(data1[p1][0].toString().equalsIgnoreCase(dpsi)){
			i1 =p1;
			break;
		 }}
		return data1[i1][data].toString();
	}


	private static String findDipsi(Sheet sheet,String SheetnameFind) throws Exception {
		Object[][] data11=FindDpsiFromSheet(sheet);
		
	
		int i11 =0;
		 for(int p1=0;p1<data11.length;p1++)
	        {
	        if(data11[p1][0].toString().equalsIgnoreCase(SheetnameFind)){
			i11 =p1;
			break;
		 }}
		return data11[i11][1].toString();
	}
	
	public static Object[][] FindDpsiFromSheet(Sheet sheet){
		Object[][] data = null;
		data = new Object[sheet.getLastRowNum()][sheet.getRow(0).getLastCellNum()];
		for (int i = 0; i < sheet.getLastRowNum(); i++) {
			for (int k = 0; k < sheet.getRow(0).getLastCellNum(); k++) {
				if (sheet.getRow(i + 1).getCell(k)!= null) {
				data[i][k] = sheet.getRow(i + 1).getCell(k).toString();
						
			}

				}
		}
		return data;
	}
	private static int CountIssues(List<String> list ,String str) {
		return Collections.frequency(list, str);
	}
	
	private static String dataFromInt(String str){
		return String.valueOf((int) Math.round(Double.parseDouble((str))));
	} 
	private static int dataFromString(String str){
		return (int) Math.round(Double.parseDouble((str)));
	} 
	
	
	@SuppressWarnings({"deprecation" })
	private static List<String> findDifferencesinExcelCell(String sheetname,String excelFilePath) {	
		List<String>SheetNamea=new ArrayList<String>();
		Xls_ReaderAsia testData = new Xls_ReaderAsia(excelFilePath);
		try {		
		Object[][] excelobj = testData.getTestData(sheetname,excelFilePath);		    	
    	int row=testData.getRowCount(sheetname);
    	
    	if (testData.isSheetExist(sheetname)) {
			if(!testData.getcolumnName(sheetname,"Diff (I)",excelFilePath)){
			testData.addColumn(sheetname, "Diff (I)");}
			if(!testData.getcolumnName(sheetname,"Diff (O)",excelFilePath)){
			testData.addColumn(sheetname, "Diff (O)");}
	//		if(!testData.getcolumnName(sheetname,"Status",excelFilePath)){
	//		testData.addColumn(sheetname, "Status");}			
		}
    	if(row==1){
    		SheetNamea.add(sheetname);			
    	}
    	boolean flag = false;
    	for(int i1=0;i1<row-1;i1++){
    		
    		String inputValue = excelobj[i1][testData.columnName(sheetname,"Input",excelFilePath)].toString();
    		String outputValue = excelobj[i1][testData.columnName(sheetname,"Output",excelFilePath)].toString();    		
    		if(!inputValue.isEmpty()&&!outputValue.isEmpty()){ 	
    				flag = true;
    		List<List<String>> resultSetToExcel=updateResults(outputValue,inputValue);
    		if (resultSetToExcel != null){    			
    			if(!resultSetToExcel.get(0).isEmpty()){
    				try{
				testData.setCellData(excelFilePath,sheetname, "Diff (I)", i1 + 2,String.join(", ", resultSetToExcel.get(0)));
			//	testData.setCellData(sheetname, "Status", i1 + 2,"Need to Report to the Dev or Deffered");
    			}
    			catch(Exception e){
    				e.printStackTrace();
    				//FileUtils.writeStringToFile(new File("./target/DifffolderTxtResults/"+excelFilePath+sheetname.replaceAll("\\s+","")+".txt"), String.join(", ", resultSetToExcel.get(0)));
    				FileUtils.writeStringToFile(new File("./Results/SummaryResults/IssueSheet/"+excelFilePath.replaceAll("\\s+","")+sheetname.replaceAll("\\s+","")+".txt"), "Issue in this sheet To compute the difference");  				
    			}	
    			
    			}
    			try{

    			if(!resultSetToExcel.get(1).isEmpty()){
    				testData.setCellData(excelFilePath,sheetname, "Diff (O)", i1 + 2,String.join(", ", resultSetToExcel.get(1)));
    		//		testData.setCellData(sheetname, "Status", i1 + 2,"Need to Report to the Dev or Deffered");
    			}}
    			catch(Exception e){
    			e.printStackTrace();
    			FileUtils.writeStringToFile(new File("./Results/SummaryResults/IssueSheet/"+excelFilePath+sheetname+".txt"), "Issue in this sheet To compute the difference");
    			}
    			
    			
    			if(resultSetToExcel.get(0).isEmpty() && resultSetToExcel.get(1).isEmpty()){
    			//	if(!findDiffIndexes(inputValue,outputValue).isEmpty()){
    			//	testData.setCellData(sheetname, "Status", i1 + 2,findDiffIndexes(inputValue,outputValue).toString());
    			
    				}
    			}
    		}
    		else{    			
				if(flag){}
    			else{
    			SheetNamea.add(sheetname);			
    			}}
    	}	
		
		}
    catch(Exception e){
    	e.printStackTrace();

    }		
		return findAlluniqueElementList(SheetNamea);
	}
	
	public static List<String> findAlluniqueElementList(List<String> list) {
      	 List<String> listWithoutDuplicates = Lists.newArrayList(Sets.newHashSet(list));
			return listWithoutDuplicates;
      	}
	
	public static List<List<String>> updateResults(String first,String second) {
		List<List<String>> dataUpdateToExcelList = null;		
		dataUpdateToExcelList = new ArrayList<List<String>>();
		String strArray[] = first.split(" ");
		String strArray2[] = second.split(" ");
		List<String> listA = Arrays.asList(strArray);
		List<String> listB = Arrays.asList(strArray2);
		List<String> firstCell=findAllDifferences(listA,listB);
		List<String> secondCell=findAllDifferences(listB,listA);
	//	System.out.println("the count space of first string is "+countsapce(first));
	//	System.out.println("the count space of second string is "+countsapce(second));
		dataUpdateToExcelList.add(firstCell);
		dataUpdateToExcelList.add(secondCell);
		return dataUpdateToExcelList;
}
	private static List<String> findAllDifferences(List<String> listA, List<String> listB) {
		List<String> result = listB.stream()
                .filter(not(new HashSet<>(listA)::contains))
                .collect(Collectors.toList());
		return result;
		}
	private static <T> Predicate<T> not(Predicate<T> predicate) {
	    return predicate.negate();
	}
	
	
	public static List<String> findDiffIndexes(String s1, String s2 ) {
	    List<String> indexes = new ArrayList<String>();
	    for( int i = 0; i < s1.length() && i < s2.length(); i++ ) {
	        if(s1.charAt(i) != s2.charAt(i)) {
	            indexes.add("String are same but index is not same");
	            break;
	        }
	    }
	    return indexes;
	}
	public void listFilesForFolder(final File folder) {
	    for (final File fileEntry : folder.listFiles()) {
	        if (fileEntry.isDirectory()) {
	            listFilesForFolder(fileEntry);
	        } else {
	            System.out.println(fileEntry.getName());
	        }
	    }
	}
	 public static boolean isNull(Object obj) {
	     return obj == null;
	 }
}