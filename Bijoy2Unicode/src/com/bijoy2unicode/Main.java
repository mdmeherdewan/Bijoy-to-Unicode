package com.bijoy2unicode;

import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import javax.script.Invocable;
import javax.script.ScriptEngine;
import javax.script.ScriptEngineManager;
import javax.script.ScriptException;
import javax.swing.JButton;
import javax.swing.JFileChooser;
import javax.swing.JFrame;
import javax.swing.JOptionPane;

import java.io.*;
import java.math.BigDecimal;
import java.math.MathContext;
import java.nio.file.Path;

public class Main {
	static File currDir = new File(".");
    static String path = currDir.getAbsolutePath();
    static String fileFolder = path.substring(0, path.length() - 1);
	
	public File getFilePathAndName(){
		 	File selectedFile = null;
	        JFileChooser jFileChooser = new JFileChooser();
	        jFileChooser.setCurrentDirectory(new File("/User/alvinreyes"));
	         
	        int result = jFileChooser.showOpenDialog(new JFrame());
	        
	        if (result == JFileChooser.APPROVE_OPTION) {
	            selectedFile = jFileChooser.getSelectedFile();
	        }
	        
	        return selectedFile;
	    }

    public static void main(String[] args) throws IOException {
    	Main m = new Main();
    	
    	File file = m.getFilePathAndName();
    	String filePath = file.getAbsolutePath();
    	String filename=file.getName(); 
    	
    	int totalSheet = 0;
    	Workbook workbook = new XSSFWorkbook();
    	System.out.println("Selected file directory : " + filePath+"\nPlease wait.......");
        FileInputStream fis=new FileInputStream(new File(filePath));
        
//      String colonDelimited = "meinsdfsdf.xls"; 
//		String[] numbers = colonDelimited.split("\\.");
//		System.out.println(numbers[numbers.length-1]); 
        
		System.out.println(fis); 
        
        XSSFWorkbook wb=new XSSFWorkbook(fis); // for input file xlsx format
       // HSSFWorkbook wb=new HSSFWorkbook(fis); // for input file xls format 
        totalSheet = wb.getNumberOfSheets();
        
        for (int sheetNo = 0; sheetNo < totalSheet; sheetNo++) {
		
	        XSSFSheet sheet=wb.getSheetAt(sheetNo); // for input file xlsx format
	        
	       // HSSFSheet sheet=wb.getSheetAt(0);// for input file xls format 
	        FormulaEvaluator formulaEvaluator=wb.getCreationHelper().createFormulaEvaluator();
	        String cellValue="0";
	        int i = 0;
	        String sheetNames = wb.getSheetName(sheetNo);
	        Sheet createSheet = workbook.createSheet(sheetNames);
	
	        for(Row row: sheet)
	        {
	            int j = 0;
	            Row createRow = createSheet.createRow(i);
	            for (Cell cell : row)
	            {
	                Cell createCell = createRow.createCell(j);
	                switch (cell.getCellType()) {
	                    case Cell.CELL_TYPE_STRING:
	                        cellValue = cell.getStringCellValue();
	                        break;
	
	                    case Cell.CELL_TYPE_FORMULA:
	                        cellValue = cell.getCellFormula();
	                        break;
	
	                    case Cell.CELL_TYPE_NUMERIC:
	                        if (DateUtil.isCellDateFormatted(cell)) {
	                            cellValue = cell.getDateCellValue().toString();
	                        } else {
	                            BigDecimal b = new BigDecimal(cell.getNumericCellValue(), MathContext.DECIMAL64);
	                            cellValue = String.valueOf(b);
	                        }
	                        break;
	
	                    case Cell.CELL_TYPE_BLANK:
	                        cellValue = "";
	                        break;
	
	                    case Cell.CELL_TYPE_BOOLEAN:
	                        cellValue = Boolean.toString(cell.getBooleanCellValue());
	                        break;
	                }
	                createCell.setCellValue(unicode(cellValue));
	                j++;
	            }
	            i++;
	        }
	    }
       
        
        String home = System.getProperty("user.home");
        String fileLocation = home+"\\Downloads\\"+filename;

        FileOutputStream outputStream = new FileOutputStream(fileLocation);
        workbook.write(outputStream);
        workbook.close();
        
        System.out.println("Made in Unicode. File Directory is bellow: \n"+fileLocation);
    }

    public static String unicode(String data){
        ScriptEngine engine = new ScriptEngineManager().getEngineByName("nashorn");
        try {
            Characters characters = new Characters();
            engine.eval(new FileReader(fileFolder+"js\\converter.js"));
            Invocable invocable = (Invocable) engine;
            String convertedFrom = "bijoy";
            for(int i=0; i<characters.listOfCharacters.length; i++){
                boolean check = data.contains(characters.listOfCharacters[i]);
                if(check==true){
                    String result;
                    result = (String) invocable.invokeFunction("ConvertToUnicode", convertedFrom, data);
                    data = result;
                    break;
                }
            }
            return data;
        }
        catch (FileNotFoundException | NoSuchMethodException | ScriptException e) {
            return data;
        }
        catch (IOException e) {
            return data;
        }
    }
 
}
