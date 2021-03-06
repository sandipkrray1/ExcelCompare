package testRunners;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Collections;
import java.util.HashMap;
import java.util.HashSet;
import java.util.List;
import java.util.Map;
import java.util.Properties;
import java.util.function.Predicate;
import java.util.stream.Collectors;
import org.apache.commons.io.FileUtils;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.google.common.base.Strings;
import com.google.common.collect.Lists;
import com.google.common.collect.Sets;

public class ExcelDifferenceupdate {
	 public static void main(String[] args) throws Exception {	         	        
			List<String>Excelnames=new ArrayList<String>();
    		List<String>Sheetnames=new ArrayList<String>();
    		List<String>Status=new ArrayList<String>();
    		List<String>DpsiValues=new ArrayList<String>();
    		List<String>CollectionValue=new ArrayList<String>();
    		List<String>SourceNameValues=new ArrayList<String>();
    		List<String>DITHLCTValues=new ArrayList<String>();
    		List<String>ContentTypeValues=new ArrayList<String>();
    		List<String>GDSVolumeValues=new ArrayList<String>();
    		List<String>PcsiValues=new ArrayList<String>();
    		List<String>SourceBundleNameValues=new ArrayList<String>();
    		List<String>BundleIDValues=new ArrayList<String>();
    		List<String>ResearchPSLValues=new ArrayList<String>();
    		List<String>ICSNonICSValues=new ArrayList<String>();
    		List<String>ExcelFileNameOnly=new ArrayList<String>();
    		List<String>DefectID=new ArrayList<String>();
    		List<String>DefectTitleValues=new ArrayList<String>();
    		
    			long startTime = System.currentTimeMillis();
		 		InputStream input = new FileInputStream("./src/test/resources/propertyFiles/config.properties");
				Properties prop = new Properties(); 
				prop.load(input);
				String SummaryResultUpdated=prop.getProperty("MasterGDSExcel");
				String summerysheetName=prop.getProperty("MasterGdsSheet");
				String DpsiDataMappingExcel=prop.getProperty("DpsiDataMappingExcel");
				//String dipsiMapping=prop.getProperty("dipsiMapping");//CSV report sheet name
				String batchMapping=prop.getProperty("batchMapping");//Master sheet name
				String LoadDate=prop.getProperty("LoadDate");
				String WaveData=prop.getProperty("WaveData");
				String TestDate=prop.getProperty("TestDate");
				//Sheet dpsisheet=GetSheetFromFile(dipsiMapping,DpsiDataMappingExcel);
				Sheet otherdatausingDpsisheet=GetSheetFromFile(batchMapping,DpsiDataMappingExcel);
				Xls_ReaderAsia testData1 = new Xls_ReaderAsia(SummaryResultUpdated);
				testData1.removeEmptyRows(summerysheetName);
        		testData1.removeRowsFromSheet(summerysheetName,SummaryResultUpdated);  
				
				
			//	String filesAre=prop.getProperty("filesAre");
				
				String filesis=prop.getProperty("filesAre");
				
				File folderis = new File(filesis);
				File[] listOfFilesis = folderis.listFiles();
				for (File fileis : listOfFilesis) {
					String DPSIName=fileis.getName().trim().toUpperCase();
					//System.out.println("DPSI is: "+DPSIName);
	        	    if (!fileis.isFile()){
	        	    	String filesis1=fileis.getPath();
	        	    	File folderis1 = new File(filesis1);
	    				File[] listOfFilesis1 = folderis1.listFiles();
	    				for (File fileis1 : listOfFilesis1) {
		        	    	String bundleName=fileis1.getName().trim().toUpperCase();
							//System.out.println("Bundle is: "+bundleName);
	    	        	    if (!fileis1.isFile()){
	    	        	    	String filesAre=fileis1.getPath()+"/";
				System.out.println("Execution is going to start at"+filesAre.substring(0, filesAre.length() - 1));
				File folder = new File(filesAre);
					    		
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
		        		sheetNames=testData1.SheetsName(filesAre+excelFiles.get(i));     		
			        	for (int j=0; j<sheetNames.size(); j++) {
			        		List<String>Status1=new ArrayList<String>();
			        		Status1=findDifferencesinExcelCell(sheetNames.get(j),filesAre+excelFiles.get(i));
			        		Sheetnames.add(sheetNames.get(j));
			        		if(!Status1.isEmpty()){		
			        			Xls_ReaderAsia testDatanew = new Xls_ReaderAsia(filesAre+excelFiles.get(i));
			        			if(Sheetnames.contains(sheetNames.get(j))){
			        				testDatanew.removeSheet(sheetNames.get(j),filesAre+excelFiles.get(i));
			        				Status.add("Pass");
			        				if(testDatanew.getFileSize(filesAre+excelFiles.get(i))==0){
			        					//System.out.println("File is going to delete");
			        					testDatanew.DelteZeroSizeFile(filesAre+excelFiles.get(i));
			        				}
			        			}		        		
			        			}
			        		else{
			        			//testData1.removeSheet(sheetNames.get(j),filesAre+excelFiles.get(i));
			        			Status.add("Fail");
			        			}

			        		
			        		//String Dpsi=findDipsi(dpsisheet,sheetNames.get(j));
			        		String Dpsi=DPSIName;
			        		//String bundle=dataFromInt(findBundle(dpsisheet,sheetNames.get(j)));
			        		String bundle=bundleName;
		    				DpsiValues.add(Dpsi);
		    				String collectionNameFrm=DpsiOtherMatchData(otherdatausingDpsisheet,Dpsi,bundle,2);
		    				CollectionValue.add(collectionNameFrm);
		    				String sourceNameFrom=DpsiOtherMatchData(otherdatausingDpsisheet,Dpsi,bundle,3);		    				
		    				SourceNameValues.add(sourceNameFrom);
		    				String DITHLCTFrom=DpsiOtherMatchData(otherdatausingDpsisheet,Dpsi,bundle,4);
		    				DITHLCTValues.add(DITHLCTFrom);
		    				String contentTypeFrom=DpsiOtherMatchData(otherdatausingDpsisheet,Dpsi,bundle,5);
		    				ContentTypeValues.add(contentTypeFrom);
		    				String GDSVolumeDPSIWiseFrom=DpsiOtherMatchData(otherdatausingDpsisheet,Dpsi,bundle,6);
		    	    		GDSVolumeValues.add(dataFromInt(GDSVolumeDPSIWiseFrom));
		    	    		String PCSIFrom=DpsiOtherMatchData(otherdatausingDpsisheet,Dpsi,bundle,7);
		    	    		PcsiValues.add(dataFromInt(PCSIFrom));
		    	    		String SourceBundleNameFrom=DpsiOtherMatchData(otherdatausingDpsisheet,Dpsi,bundle,8);
		    	    		SourceBundleNameValues.add(SourceBundleNameFrom);
		    	    		String BundleIDFrom=DpsiOtherMatchData(otherdatausingDpsisheet,Dpsi,bundle,9);
		    	    		BundleIDValues.add(dataFromInt(BundleIDFrom));
		    	    		String ResearchPSL=DpsiOtherMatchData(otherdatausingDpsisheet,Dpsi,bundle,10);
		    	    		ResearchPSLValues.add(ResearchPSL);
		    	    		String ICSNonICS=DpsiOtherMatchData(otherdatausingDpsisheet,Dpsi,bundle,11);
		    	    		ICSNonICSValues.add(ICSNonICS);
		    	    		String defectID=DpsiOtherMatchData(otherdatausingDpsisheet,Dpsi,bundle,12);
		    	    		DefectID.add(defectID);
		    	    		String defectTitle=DpsiOtherMatchData(otherdatausingDpsisheet,Dpsi,bundle,13);
		    	    		DefectTitleValues.add(defectTitle);
		    	    			    		
		    	    		//String ExcelFileNameupdated=Dpsi+"_"+dataFromInt(PCSIFrom)+"_"+dataFromInt(BundleIDFrom)+"_"+excelFiles.get(i);
		    	    		Excelnames.add(Dpsi+"_"+dataFromInt(PCSIFrom)+"_"+dataFromInt(BundleIDFrom)+"_"+excelFiles.get(i));
		    	    		ExcelFileNameOnly.add(excelFiles.get(i));
			        	}
			        	long endTime1 = System.currentTimeMillis();
			        	System.out.println("That took " + (endTime1 - startTime1)/1000 + "seconds in Excel--> "+excelFiles.get(i));    
        		}
        		
        		
        		//Excelnames.forEach(System.out::println);
	    		//	Sheetnames.forEach(System.out::println);
	    		//	Status.forEach(System.out::println);

        	//	testData1.removeEmptyRows(summerysheetName);
        	//	testData1.removeRowsFromSheet(summerysheetName,SummaryResultUpdated);        		

        		FileInputStream fis = new FileInputStream(SummaryResultUpdated);
	        	Workbook wb = WorkbookFactory.create(fis);
	        	Sheet sheet = wb.getSheet(summerysheetName);
	        	int rowCount = 0;	        	
	        	
	    		for (int p = 0; p < Excelnames.size(); p++) {   		
	    			
	    			Row row = sheet.createRow(++rowCount);
	    		
	    		//System.out.println("Last row no is"+newrow);
	    			//iterating sr.no number of columns
	    			for (int c=0;c < 1; c++ )
	    			{
	    			//Sr.no
	    			Cell cell0 = row.createCell(c);
	    			cell0.setCellValue(p+1);
	    			//dpsi values 			
	    			Cell cell1 = row.createCell(c+1);
	    			cell1.setCellValue(DpsiValues.get(p));
	    			//WaveInfo
	    			Cell cell2 = row.createCell(c+2);	    			
	    			cell2.setCellValue(WaveData);
	    			//Research / PSL
	    			Cell cell3 = row.createCell(c+3);	    			
	    			cell3.setCellValue(ResearchPSLValues.get(p));
	    			//ICS- Non Ics Values
	    			Cell cell4 = row.createCell(c+4);	    			
	    			cell4.setCellValue(ICSNonICSValues.get(p));
	    			//Collection name values
	    			Cell cell5 = row.createCell(c+5);
	    			cell5.setCellValue(CollectionValue.get(p));
	    			//SourceName Values
	    			Cell cell6 = row.createCell(c+6);
	    			cell6.setCellValue(SourceNameValues.get(p));
	    			//DITHLCT values 
	    			Cell cell7 = row.createCell(c+7);
	    			cell7.setCellValue(DITHLCTValues.get(p));
	    			//ContentType Values 
	    			Cell cell8 = row.createCell(c+8);
	    			cell8.setCellValue(ContentTypeValues.get(p));
	    			//GDSVolume Values 
	    			Cell cell9 = row.createCell(c+9);
	    			cell9.setCellValue(dataFromString(GDSVolumeValues.get(p)));    			
	    			//Pcsi Values
	    			Cell cell10 = row.createCell(c+10);
	    			cell10.setCellValue(dataFromString(PcsiValues.get(p)));    			
	    			// SourceBundleNameValues
	    			Cell cell11 = row.createCell(c+11);
	    			cell11.setCellValue(SourceBundleNameValues.get(p));    			
	    			//	BundleIDValues
	    			Cell cell12 = row.createCell(c+12);
	    			cell12.setCellValue(dataFromString(BundleIDValues.get(p)));    			
	    			//Sheet name 
	    			Cell cell13 = row.createCell(c+13);
	    			cell13.setCellValue(Sheetnames.get(p));
	    			//Excel file name 
	    			Cell cell14 = row.createCell(c+14);
	    			cell14.setCellValue(Excelnames.get(p));
	    			//Status  
	    			Cell cell15 = row.createCell(c+15);	    			
	    			cell15.setCellValue(Status.get(p));
	    			//LoadDate
	    			Cell cell16 = row.createCell(c+16);	    			
	    			cell16.setCellValue(LoadDate);	    			
	    			//TestDate
	    			Cell cell17 = row.createCell(c+17);	    			
	    			cell17.setCellValue(TestDate);
	    			
	    		//	DefectID values 
	    			Cell cell18 = row.createCell(c+18);	    			
	    			cell18.setCellValue(DefectID.get(p));
	    		
		    		//	DefectTitleValues values 
	    			Cell cell19 = row.createCell(c+19);	    			
	    			cell19.setCellValue(DefectTitleValues.get(p));

	    			
	    			//ExcelFileNameonly
	    		//	Cell cell20 = row.createCell(c+22);	    			
	    		//	cell20.setCellValue(ExcelFileNameOnly.get(p));
	    			}
	    		}	    		
		    		
	    		FileOutputStream fos = new FileOutputStream(SummaryResultUpdated);
	    		//Write this workbook to an Outputstream.
	    		wb.write(fos);
	    		fos.flush();
	    		fos.close();
	    		wb.close();
	         	

	                	   	
	  }
        	    }
        	    }}
	        	long endTime = System.currentTimeMillis();
				System.out.println("That took " + (endTime - startTime)/1000 + "seconds");
				fileDelteFrom();
				System.out.println("Data is updated");			
	 
	 }
	 
	 public static void fileDelteFrom() throws Exception{
	 		InputStream input = new FileInputStream("./src/test/resources/propertyFiles/config.properties");
			Properties prop = new Properties(); 
			prop.load(input);
			String filesis=prop.getProperty("filesAre");
			File folderis = new File(filesis);
			File[] listOfFilesis = folderis.listFiles();
			for (File fileis : listOfFilesis) {
				if (!fileis.isFile()){
  	    	String filesis1=fileis.getPath();
  	    	File folderis1 = new File(filesis1);
				File[] listOfFilesis1 = folderis1.listFiles();
				for (File fileis1 : listOfFilesis1) {
	        	    if (!fileis1.isFile()){
	        	    	String filesAre=fileis1.getPath()+"/";
			System.out.println("Going to delete empty file at "+filesAre.substring(0, filesAre.length() - 1));
			File folder = new File(filesAre);
			File[] listOfFiles = folder.listFiles();
			for (File file : listOfFiles) {
  	    if (file.isFile()) {
  	    if(file.getName().contains(".xlsx")){
  	    	if(file.length()==0){
  	    		System.gc();
  	    		file.deleteOnExit();
  	    	}
  	    }}}}}}}
	 }

	@SuppressWarnings("resource")
	public static Sheet GetSheetFromFile(String excelSheet,String ExcelFile) throws Exception{
		FileInputStream fis = new FileInputStream(ExcelFile);
		XSSFWorkbook  workbook = new XSSFWorkbook(fis);
	    //XSSFSheet sheet = workbook.getSheetAt(0);
	    XSSFSheet sheet = workbook.getSheet(excelSheet);	    
	    return sheet;
	}
	
	private static String DpsiOtherMatchData(Sheet sheet,String dpsi,String bundle,int data) throws Exception {
		Object[][] data1=FindDpsiFromSheet(sheet);
		boolean flag = false;
		int i1 =0;
		 for(int p1=0;p1<data1.length;p1++)
	        {
			 if(data1[p1][1].toString().equalsIgnoreCase(dpsi)&&dataFromInt(data1[p1][9].toString()).equalsIgnoreCase(bundle)){
				 i1 =p1;
				 flag = true;
				 break;
			 }
	     }
		 if(flag==false)
			 System.out.println("DPSI("+dpsi+")/BundleID("+bundle+") does not matched with Mastersheet details.");

		 return data1[i1][data].toString();
	}


	private static String findDipsi(Sheet sheet,String SheetnameFind) throws Exception {
		Object[][] data11=FindDpsiFromSheet(sheet);
		int i11 =0;
		 for(int p1=0;p1<data11.length;p1++)
	        {
	        if(data11[p1][5].toString().equalsIgnoreCase(SheetnameFind)){
			i11 =p1;
			break;
		 }}
		return data11[i11][2].toString();
	}
	private static String findBundle(Sheet sheet,String SheetnameFind) throws Exception {
		Object[][] data11=FindDpsiFromSheet(sheet);
		int i11 =0;
		 for(int p1=0;p1<data11.length;p1++)
	        {
	        if(data11[p1][5].toString().equalsIgnoreCase(SheetnameFind)){
			i11 =p1;
			break;
		 }}
		return data11[i11][3].toString();
	}
	public static Object[][] FindDpsiFromSheet(Sheet sheet){
		Object[][] data = null;
		data = new Object[sheet.getLastRowNum()][sheet.getRow(0).getLastCellNum()];
		for (int i = 0; i < sheet.getLastRowNum(); i++) {
			for (int k = 0; k < sheet.getRow(0).getLastCellNum(); k++) {
				if (sheet.getRow(i + 1).getCell(k)!= null) {
				data[i][k] = sheet.getRow(i + 1).getCell(k).toString();
			}}
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
		Xls_ReaderAsia testData1 = new Xls_ReaderAsia(excelFilePath);
		try {		
		//Object[][] excelobj = testData.getTestData(sheetname,excelFilePath);		    	
    	int row=testData1.getRowCount(sheetname);
		deletusingRecusion(sheetname,excelFilePath,0);
    	if (testData1.isSheetExist(sheetname)) {
			if(!testData1.getcolumnName(sheetname,"Diff (I)",excelFilePath)){
			testData1.addColumn(sheetname, "Diff (I)");}
			if(!testData1.getcolumnName(sheetname,"Diff (O)",excelFilePath)){
			testData1.addColumn(sheetname, "Diff (O)");}

			if(!testData1.getcolumnName(sheetname,"I/P Scan",excelFilePath)){
				testData1.addColumn(sheetname, "I/P Scan");}
			if(!testData1.getcolumnName(sheetname,"O/P Scan",excelFilePath)){
				testData1.addColumn(sheetname, "O/P Scan");}
	//		if(!testData.getcolumnName(sheetname,"Status",excelFilePath)){
	//		testData.addColumn(sheetname, "Status");}			
		}

		if(testData1.getRowCount(sheetname)==1){
			SheetNamea.add(sheetname);			
    	}
		else{

    	boolean flag = false;
    	for(int i1=0;i1<Xls_ReaderAsia.getRowsCount(excelFilePath,sheetname)-1;i1++){
   		 Xls_ReaderAsia testData = new Xls_ReaderAsia(excelFilePath);
			 Object[][] excelobj = testData.getTestData(sheetname,excelFilePath);	
   		String inputValue = excelobj[i1][testData.columnName(sheetname,"Input",excelFilePath)].toString();
   		String outputValue = excelobj[i1][testData.columnName(sheetname,"Output",excelFilePath)].toString();    		
   		Map<Integer, String> thirdCellList=new HashMap<Integer, String>();
   		Map<Integer, String> fourthCellList=new HashMap<Integer, String>();
   		thirdCellList=Xls_ReaderAsia.getTestDatasummary(excelFilePath,sheetname,3);
   		fourthCellList=Xls_ReaderAsia.getTestDatasummary(excelFilePath,sheetname,4);
    		
    		if(!inputValue.isEmpty()&&!outputValue.isEmpty()){ 	
    				flag = true;
        			List<List<String>> resultSetToExcel=updateResults(outputValue,inputValue,thirdCellList,fourthCellList);
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
    			
    			
    		//	if(resultSetToExcel.get(0).isEmpty() && resultSetToExcel.get(1).isEmpty()){
    			//	if(!findDiffIndexes(inputValue,outputValue).isEmpty()){
    			//	testData.setCellData(sheetname, "Status", i1 + 2,findDiffIndexes(inputValue,outputValue).toString());
    			
    			//	}
    			if(flag){
					if (SheetNamea.contains(sheetname)) {
					SheetNamea.remove(0);
					}
				}
    		
    		}
    		}
    		else{    			
				if(flag){
					if (SheetNamea.contains(sheetname)) {
					SheetNamea.remove(0);
					}
				}
    			else{
    				if (!SheetNamea.contains(sheetname)) {
    					SheetNamea.add(sheetname);
    				}
    			}}
    		
    		List<List<String>> resultSetToExcel=updateResults(outputValue,inputValue,thirdCellList,fourthCellList);
    		if (resultSetToExcel != null){    			
    			if(!resultSetToExcel.get(2).isEmpty()){
    				try{
        				testData.setCellData(excelFilePath,sheetname, "I/P Scan", i1 + 2,String.join(", ", resultSetToExcel.get(2)));
        			
            			}
            			catch(Exception e){
            				e.printStackTrace();
            			
            			}
    				try{
        				testData.setCellData(excelFilePath,sheetname, "O/P Scan", i1 + 2,String.join(", ", resultSetToExcel.get(3)));
        			
            			}
            			catch(Exception e){
            				e.printStackTrace();
            			
            			}
    			}	

    	}	

	}
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
	 
	 
	 public static void deletusingRecusion(String sheetname,String excelFilePath,int t) throws Exception{	
			

			
		 for(int i2=0;i2<(Xls_ReaderAsia.getRowsCount(excelFilePath,sheetname)-1);i2++){	
			 Xls_ReaderAsia testData = new Xls_ReaderAsia(excelFilePath);
			 Object[][] excelobj = testData.getTestData(sheetname,excelFilePath);						
			 String difftype=excelobj[i2][testData.columnName(sheetname,"Diff Type",excelFilePath)].toString();
		    	String input=excelobj[i2][testData.columnName(sheetname,"Input",excelFilePath)].toString();
		    	String output=excelobj[i2][testData.columnName(sheetname,"Output",excelFilePath)].toString();
		    	if(!difftype.equalsIgnoreCase("CHANGE")){
		    		if(input.equalsIgnoreCase("<p>")&&output.isEmpty()||input.equalsIgnoreCase("</p>")&&output.isEmpty()||output.equalsIgnoreCase("<p>")&&input.isEmpty()||output.equalsIgnoreCase("</p>")&&input.isEmpty()){
		    			//rows.add(i2);
		    			// System.out.println("Going to delete-->"+(i2+1));	
		    			testData.deleteRow(excelFilePath,sheetname,(i2+1));
		    			i2=i2-1;
		    		}}
		    	}
		        }
		public static List<List<String>> updateResults(String first,String second, Map<Integer, String> thirdCellList, Map<Integer, String> fourthCellList) throws Exception {
			List<List<String>> dataUpdateToExcelList = null;		
			dataUpdateToExcelList = new ArrayList<List<String>>();
			String strArray[] = first.split(" ");
			String strArray2[] = second.split(" ");
			List<String> listA = Arrays.asList(strArray);
			List<String> listB = Arrays.asList(strArray2);
			List<String> firstCell=findAllDifferences(listA,listB);
			List<String> secondCell=findAllDifferences(listB,listA);
			
			String third=second;	
			List<String>fourth=new ArrayList<String>();
			List<String>thirdAA=new ArrayList<String>();
			//For Last node update
			String four=first;	
			List<String>thirdlast=new ArrayList<String>();
			List<String>fourthAA=new ArrayList<String>();
			fourthCellList.entrySet().forEach(entry -> {
				fourth.add(entry.getValue());
			});
			thirdCellList.entrySet().forEach(entry -> {
				thirdlast.add(entry.getValue());
			});
			
			//for 3rd node compare logic 
			if(!Strings.isNullOrEmpty(third)){
			int index=findIndexofData(fourth,third);
			if(index>0){
				thirdAA.add("Found at->"+String.valueOf(index+1));
			}
			else{
				thirdAA.add("not Found");			
			}}
			else{
				thirdAA.add("N/A");
			}
		//for four node compare logic 
			if(!Strings.isNullOrEmpty(four)){
				int index=findIndexofData(thirdlast,four);
				if(index>0){
					fourthAA.add("Found at->"+String.valueOf(index+1));
				}
				else{
					fourthAA.add("not Found");			
				}}
				else{
					fourthAA.add("N/A");
				}
			dataUpdateToExcelList.add(firstCell);
			dataUpdateToExcelList.add(secondCell);
			dataUpdateToExcelList.add(thirdAA);
			dataUpdateToExcelList.add(fourthAA);
			return dataUpdateToExcelList;
	}
	 
		public static int findIndexofData(List<String> a, String s){
			 int index=0;
			 for(int i=0;i<a.size();i++){
				 if(a.get(i).contains(s)){
					 index=i+1;
					 break;
				 }
			 }
			 if(index==0){
				 index=-1;
			 }
			return index;	 
			 
		 }
		
}