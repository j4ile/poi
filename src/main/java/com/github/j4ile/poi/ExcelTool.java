package com.github.j4ile.poi;

import java.awt.Color;
import java.io.BufferedReader;
import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStreamReader;
import java.io.PrintStream;
import java.time.LocalDate;
import java.time.temporal.ChronoUnit;
import java.util.ArrayList;
import java.util.List;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Hyperlink;
import org.apache.poi.ss.usermodel.RichTextString;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.Row.MissingCellPolicy;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFCreationHelper;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFFormulaEvaluator;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;



public class ExcelTool {

    private static boolean skipTotals;
    final XSSFWorkbook workbook;
    XSSFSheet sheet;
    XSSFCreationHelper createHelper;

    int columnWriterRow = 0;
    int columnWriterCol = 0;
    
    public static void main (String [] args) throws IOException{
        if(args.length==0){
            System.err.println("Usage: <file> -columns <columns> -csvFile <name> -skiphead -skiptotals -nocommas");
            return;
        }
        String xlsxFile = args[0];
        int sheetNo = 0;
        int index=1;
        String csvFile = null;
        List<Integer> columns = null;
        boolean skipHead=false;
        boolean nocommas=false;
        while(index<args.length){
            if(args[index].equalsIgnoreCase("-sheet")){
                sheetNo = Integer.parseInt(args[++index])-1;
            }
            else if(args[index].equalsIgnoreCase("-skiphead")){
                skipHead=true;
            }
            else if(args[index].equalsIgnoreCase("-skiptotals")){
                skipTotals=true;
            }
            else if(args[index].equalsIgnoreCase("-nocommas")){
                nocommas=true;
            }
            else if(args[index].equalsIgnoreCase("-csvFile")){
                csvFile = args[++index];
            }
            else if(args[index].equalsIgnoreCase("-columns")){
                columns = new ArrayList<Integer>();
                for(final String s:args[++index].toUpperCase().split(",")){
                    if(s!=null && s.length()>0){
                        try{
                            columns.add(Integer.parseInt(s)-1); //Add numerics
                        }catch(Exception e){
                        	columns.add(convertColumnLetterToNumber(s));
                        }
                    }
                }
            }
            ++index;
        }
        System.out.println(columns);
        
        final PrintStream out = csvFile==null?null:new PrintStream(new FileOutputStream(csvFile),
                true, "UTF-8");
        
        final XSSFWorkbook wb = loadWorkBookFromFileName(xlsxFile);
        
        generateCSV(wb, sheetNo, skipHead, skipTotals, nocommas, columns, out);
        
        if(out!=null){
            out.close();
            System.out.println("Output written to " + csvFile);
        }
    }
    
    /**
     * This will process the file manipulation (so I can better test it)
     * @param wb - The excel document to edit..
     * @param sheetNo - sheet number we are working in
     * @param skipHead - remove the headers from the file
     * @param skipTotals - remove the last row of the sheet
     * @param nocommas - don't include commons between rows
     * @param columns - the columns 'letters' to include
     * @param out - the output stream to write the data too. 
     * @throws IOException
     */
    public static void generateCSV(final Workbook wb
    		, final int sheetNo
    		, final boolean skipHead
    		, final boolean skipTotals
    		, final boolean nocommas
    		, final List<Integer> columns
    		, final PrintStream out) throws IOException{

        final DataFormatter formatter = new DataFormatter();
        //final byte[] bom = {(byte)0xEF, (byte)0xBB, (byte)0xBF};
        //out.write(bom);
        {
            Sheet sheet = wb.getSheetAt(sheetNo);
            for (int r = skipHead?1:0, rn = skipTotals?sheet.getLastRowNum()-1:sheet.getLastRowNum() ; r <= rn ; r++) {
                Row row = sheet.getRow(r);
                if ( row == null ) { out.println(','); continue; }
                boolean firstCell = true;
                if(columns!=null&&columns.size()>0){
                    for(int c: columns){
                        if ( ! firstCell ) {
                        	if (out != null) out.print(',');
                        }
                        firstCell= false;
                        doCell(c,row,out,nocommas,formatter);
                    }
                }else{
                    for (int c = 0, cn = row.getLastCellNum() ; c < cn ; c++) {
                        if ( ! firstCell ) {
                        	if (out != null) out.print(',');
                        }
                        firstCell= false;
                        doCell(c,row,out,nocommas,formatter);
                        
                    }
                }
                if(out!=null)
                    out.println();
            }
        }
    }
    
    /**
     * Attempts to find the File, and returns 
     * @param fileName
     * @return
     * @throws IOException
     */
    public static XSSFWorkbook loadWorkBookFromFileName(final String fileName) throws IOException {
    	final String extension = fileName.lastIndexOf('.') > 0 ? fileName.substring(fileName.lastIndexOf(".")+1) : "";
    	
    	if ("csv".equals(extension)) {
    		return importCSV(fileName);
    	} 
    	return new XSSFWorkbook(new FileInputStream(fileName));
    }
    
    
    /**
     * This will convert the file to a XSSWorkbook
     * @param csvFileName
     * @return
     * @throws IOException
     */
    private static XSSFWorkbook importCSV(final String csvFileName) throws IOException{
        final XSSFWorkbook workBook = new XSSFWorkbook();
        final XSSFSheet sheet = workBook.createSheet("sheet1");
        String currentLine=null;
        int RowNum=0;
        BufferedReader br = new BufferedReader(new InputStreamReader(new FileInputStream(csvFileName)));
        
        while ((currentLine = br.readLine()) != null) {
        	List<String> clmLst = decodeCSVRow(currentLine);
            XSSFRow currentRow=sheet.createRow(RowNum++);
            for(int i=0;i<clmLst.size();i++){
            	// remove the '"' on either side, 
                currentRow.createCell(i).setCellValue(clmLst.get(i).replaceAll("(^\"+|\"+$)", ""));
            }
        }
        br.close();
        
        /*
        // this is here to quickly test what the XML version of the imported file looks like.
        FileOutputStream fileOutputStream =  new FileOutputStream("c:\\temp\\output_Text.xlsx");
        workBook.write(fileOutputStream);
        fileOutputStream.close();
        //*/
        return workBook;
    }
    
    private static List<String> decodeCSVRow(final String rowStr) {
    	final ArrayList<String> rtnList = new ArrayList<String>();
    	String activeRowContent = rowStr; 
    	int indx = -1;
    	while ((indx = FindEndOfColumn(activeRowContent)) > 0){
    		rtnList.add(activeRowContent.substring(0, indx)); // -1 to remove the ','
    		if (activeRowContent.length() > indx ) {
    			activeRowContent = activeRowContent.substring(indx + 1); 	
    		} else  {
    			activeRowContent = "";
    		}
    	}
    	return rtnList;
    }
    
    private static int FindEndOfColumn(final String rowStr) {
    	int cmIndx = rowStr.indexOf(',');
    	// if it's a string wrapped context, find the next , after the end quote
    	if (rowStr.startsWith("\"")) {
    	    Matcher matcher = (Pattern.compile("(?!<\")\",")).matcher(rowStr);
    	    cmIndx = matcher.find() ? matcher.end() - 1 : -1; // to remove the comma from the reference...
    	} 
    	if (cmIndx == -1 && !rowStr.equals("")) {
    		return rowStr.length();
    	}
    	return cmIndx;
    }
    
    
    /**
     * This will convert a letter, into a column index (starting from a)
     * 
     * It sets the string to an uppercase letter (so you only work with the one set of ASCII codes, (lowercase letters have a higher value)) Subtracts 65 (Acsii Letter A) from your current Ascii value to get the column (starting at 0) IF there were two letters, It adds the two together (multiply by 26 on the second value to make it base 26 ~esk). 
     * 
     * @param ColumnLetter
     * @return
     */
    public static int convertColumnLetterToNumber(final String ColumnLetter) {
    	final String ckLtr = ColumnLetter.toUpperCase();
    	if(ckLtr.length()>1){
             return ((int)ckLtr.charAt(1)-65+(((int)ckLtr.charAt(0)-64)*26));  //Add alpha
        }else
        return ((int)ckLtr.charAt(0)-65);  //Add alpha
    }
    

    private static void doCell(int c,Row row ,PrintStream out,boolean nocommas,DataFormatter formatter) {
        final Cell cell = row.getCell(c, MissingCellPolicy.RETURN_BLANK_AS_NULL);
        if ( cell != null ) {
            String value = formatter.formatCellValue(cell);
            if(nocommas)
                value= value.replaceAll(",", "");
            if(out!=null)
                out.print(encodeValue(value));
        }
    }

    public ExcelTool() throws IOException {
        workbook = new XSSFWorkbook();
        sheet = workbook.createSheet();
        createHelper = workbook.getCreationHelper();
    }

    public ExcelTool(final byte[] filename) throws IOException {
        this(new String(filename).trim());
    }

    public ExcelTool(final String filename) throws IOException {
        if (filename.trim().length() > 0) {
            final FileInputStream file = new FileInputStream(new File(filename));
            try {
                workbook = new XSSFWorkbook(file);
                sheet = workbook.getSheetAt(0);
            } finally {
                file.close();
            }
        } else {
            workbook = new XSSFWorkbook();
            sheet = workbook.createSheet();
        }
        createHelper = workbook.getCreationHelper();
    }

    public ExcelTool(final XSSFWorkbook workbook) {
    	this.workbook = workbook;
    	this.sheet = this.workbook.getSheetAt(0);
    	createHelper = this.workbook.getCreationHelper();
    }
    
    /**
     * First sheet is 1;
     * 
     * @param sheetNumber
     */
    public void setSheet(final int sheetNumber) {
        if (sheetNumber > 0 && sheetNumber <= 10) {
            while (workbook.getNumberOfSheets() < sheetNumber) {
                workbook.createSheet();
            }
            this.sheet = workbook.getSheetAt(sheetNumber - 1);
        }
    }

    /**
     * Duplicate a specific sheet
     * 
     * @param sheetNumber
     */
    public void duplicateSheet(final int sheetNumber) {
        if (sheetNumber > 0 && sheetNumber <= workbook.getNumberOfSheets()) {
            this.sheet = workbook.cloneSheet(sheetNumber - 1);
        }
    }

    public void setSheetName(final int sheetIndex, final String sheetname) {
        this.workbook.setSheetName(sheetIndex - 1, sheetname);
    }

    public int getCurrentSheetIndex() {
        for (int i = 0; i < this.workbook.getNumberOfSheets(); i++) {
            if (this.sheet == this.workbook.getSheetAt(i)) {
                return i + 1;
            }
        }
        return 1;
    }

    public void setCurrentSheetName(final byte[] name) {
        this.workbook.setSheetName(getCurrentSheetIndex() - 1, new String(name).trim());
    }

    /**
     * Delete rows after
     * 
     * @param startRow
     */
    public void deleteRowsAfter(final int startRow) {
        XSSFRow currentrow = null;
        while ((currentrow = this.sheet.getRow(startRow + 1)) != null) {
            this.sheet.removeRow(currentrow);
        }
    }
    public void deleteRow(final int startRow) {
        XSSFRow currentrow = null;
        if ((currentrow = this.sheet.getRow(startRow)) != null) {
            this.sheet.removeRow(currentrow);
            if(startRow<sheet.getLastRowNum())
                sheet.shiftRows(startRow+1, sheet.getLastRowNum(), -1);
        }
    }

    public void setCellInt(final int rowNum, final int colNum, final int value) {
        getCell(rowNum, colNum).setCellValue(value);
    }
    
    public void setCellExcelDate(final int rowNum, final int colNum, final byte[] date) {
        if(date.length==0){
            getCell(rowNum, colNum).setCellValue((String)null);
            return;
        }
        try{
            LocalDate ldate1= LocalDate.parse(new String(date));
            long days = Math.abs(ChronoUnit.DAYS.between(ldate1,LocalDate.parse("1900-01-01")))+2;
            getCell(rowNum, colNum).setCellValue(days);
        }
        catch(Exception e){
            getCell(rowNum, colNum).setCellValue((String)null);
        }
    }

    public int getCellInt(final int rowNum, final int colNum) {
        return Double.valueOf(getCell(rowNum, colNum).getNumericCellValue()).intValue();
    }

    public void setCellString(final int rowNum, final int colNum, final String value) {
        getCell(rowNum, colNum).setCellValue(value);
    }

    public String getCellString(final int rowNum, final int colNum) {
        return getCell(rowNum, colNum).getStringCellValue();
    }

    public void setCellRichText(final int rowNum, final int colNum, final RichTextString richText) {
        getCell(rowNum, colNum).setCellValue(richText);
    }
    public void setSheetTab(final int tabNum) {
            workbook.setSelectedTab(tabNum - 1);
    }

    public void setCellHyperlink(final int rowNum, final int colNum, final String label, final String link) {
        final Hyperlink hyperlink = createHelper.createHyperlink(org.apache.poi.common.usermodel.HyperlinkType.URL);
        hyperlink.setAddress(link);
        hyperlink.setLabel(label);
        getCell(rowNum, colNum).setHyperlink(hyperlink);
    }

    public void setCellColor(final int rowNum, final int colNum, final int r, final int g, final int b) {
        final XSSFCellStyle style = getCell(rowNum, colNum).getCellStyle();
        style.setFillBackgroundColor(new XSSFColor(new Color(r, g, b)));
        getCell(rowNum, colNum).setCellStyle(style);
    }

    public void setCellColor(final int rowNum, final int colNum, final String colorName) {
        final Color target = getColorByName(colorName);
        if (target != null) {
            final XSSFCellStyle style = getCell(rowNum, colNum).getCellStyle();
            style.setFillBackgroundColor(new XSSFColor(target));
            getCell(rowNum, colNum).setCellStyle(style);
        }
    }
    
    public void setCellFont(final int rowNum, final int colNum, final String fontName) {
        final XSSFCellStyle currentStyle = getCell(rowNum, colNum).getCellStyle();
        final XSSFFont currentFont= currentStyle.getFont();
        currentFont.setFontName(fontName);
        currentStyle.setFont(currentFont);
        getCell(rowNum, colNum).setCellStyle(currentStyle);
    }
    public void setCellFontSize(final int rowNum, final int colNum, final int fontSize) {
        final XSSFCellStyle currentStyle = getCell(rowNum, colNum).getCellStyle();
        final XSSFFont currentFont= currentStyle.getFont();
        currentFont.setFontHeightInPoints((short)fontSize);
        currentStyle.setFont(currentFont);
        getCell(rowNum, colNum).setCellStyle(currentStyle);
    }
    public void setCellFontBold(final int rowNum, final int colNum, final boolean bold) {
        final XSSFCellStyle currentStyle = getCell(rowNum, colNum).getCellStyle();
        final XSSFFont currentFont= currentStyle.getFont();
        currentFont.setBold(bold);
        currentStyle.setFont(currentFont);
        getCell(rowNum, colNum).setCellStyle(currentStyle);
    }
    public void setCellFontItalic(final int rowNum, final int colNum, final boolean italic) {
        final XSSFCellStyle currentStyle = getCell(rowNum, colNum).getCellStyle();
        final XSSFFont currentFont= currentStyle.getFont();
        currentFont.setItalic(italic);
        currentStyle.setFont(currentFont);
        getCell(rowNum, colNum).setCellStyle(currentStyle);
    }

    public void setCellBytes(final int rowNum, final int colNum, final byte[] value) {
        getCell(rowNum, colNum).setCellValue(new String(value));
    }

    public byte[] getCellBytes(final int rowNum, final int colNum) {
        return getCell(rowNum, colNum).getStringCellValue().getBytes();
    }

    public void setCellDecimal(final int rowNum, final int colNum, final double value) {
        getCell(rowNum, colNum).setCellValue(value);
    }

    public double getCellDecimal(final int rowNum, final int colNum) {
        return getCell(rowNum, colNum).getNumericCellValue();
    }

    public void setRowValues(final int rowNum, int colNum, final Object... values) {
        for (final Object val : values) {
            if (val == null) {
                colNum++;
            } else if (val instanceof Number) {
                setCellDecimal(rowNum, colNum++, ((Number) val).doubleValue());
            } else if (val instanceof byte[]) {
                setCellBytes(rowNum, colNum++, (byte[]) val);
            } else {
                setCellString(rowNum, colNum++, val.toString());
            }
        }
    }

    public void setCellFormulaString(final int rowNum, final int colNum, final String formula) {
        getCell(rowNum, colNum).setCellFormula(formula);
    }

    public void setCellFormula(final int rowNum, final int colNum, final byte[] formula) {
        setCellFormulaString(rowNum, colNum, new String(formula));
    }

    XSSFCell getCell(final int rowNum, final int colNum) {
        XSSFRow row = sheet.getRow(rowNum);
        if (row == null) {
            row = sheet.createRow(rowNum);
        }
        XSSFCell cell = row.getCell(colNum);
        if (cell == null) {
            cell = row.createCell(colNum);
        }
        return cell;
    }

    public void save(final byte[] filename) throws IOException {
        save(new String(filename).trim());
    }

    public void save(final String filename) throws IOException {
        final FileOutputStream fos = new FileOutputStream(filename.trim());
        try {
            workbook.write(fos);
        } finally {
            fos.close();
        }
    }

    public byte[] save() throws IOException {
        final ByteArrayOutputStream fos = new ByteArrayOutputStream();
        try {
            workbook.write(fos);
            return fos.toByteArray();
        } finally {
            fos.close();
        }
    }

    public void setColumnWriterPosition(final int row, final int col) {
        columnWriterRow = row;
        columnWriterCol = col;
    }

    public void write(final Object value) {
        if (value instanceof Integer) {
            setCellInt(columnWriterRow++, columnWriterCol, (Integer) value);
        } else if (value instanceof Double) {
            setCellDecimal(columnWriterRow++, columnWriterCol, (Double) value);
        } else {
            setCellString(columnWriterRow++, columnWriterCol, value.toString());
        }
    }

    public void recalc() {
        final XSSFFormulaEvaluator evaluator = workbook.getCreationHelper().createFormulaEvaluator();
       
        evaluator.evaluateAll();
    }
    
    static Color getColorByName(final String name) {
        try {
            return (Color) Color.class.getField(name.toUpperCase()).get(null);
        } catch (final Exception e) {
            return null;
        }
    }

    public XSSFWorkbook getWorkbook() {
        return workbook;
    }

    public XSSFSheet getSheet() {
        return sheet;
    }
    
    void copyRow(int sourceRowNum, int destinationRowNum) {
        // Get the source / new row
        XSSFRow newRow = sheet.getRow(destinationRowNum);
        XSSFRow sourceRow = sheet.getRow(sourceRowNum);

        // If the row exist in destination, push down all rows by 1 else create a new row
        if (newRow != null) {
            sheet.shiftRows(destinationRowNum, sheet.getLastRowNum(), 1);
            newRow = sheet.createRow(destinationRowNum);
        } 
        newRow = sheet.createRow(destinationRowNum);
        
        // Loop through source columns to add to new row
        for (int i = 0; i < sourceRow.getLastCellNum(); i++) {
            // Grab a copy of the old/new cell
            XSSFCell oldCell = sourceRow.getCell(i);
            XSSFCell newCell = newRow.createCell(i);

            // If the old cell is null jump to next cell
            if (oldCell == null) {
                newCell = null;
                continue;
            }

            // Copy style from old cell and apply to new cell
            XSSFCellStyle newCellStyle = workbook.createCellStyle();
            newCellStyle.cloneStyleFrom(oldCell.getCellStyle());
            ;
            newCell.setCellStyle(newCellStyle);

            // If there is a cell comment, copy
            if (oldCell.getCellComment() != null) {
                newCell.setCellComment(oldCell.getCellComment());
            }

            // If there is a cell hyperlink, copy
            if (oldCell.getHyperlink() != null) {
                newCell.setHyperlink(oldCell.getHyperlink());
            }

            // Set the cell data type
            newCell.setCellType(oldCell.getCellType());

            // Set the cell data value
            switch (oldCell.getCellType()) {
                case BLANK:
                    newCell.setCellValue(oldCell.getStringCellValue());
                    break;
                case BOOLEAN:
                    newCell.setCellValue(oldCell.getBooleanCellValue());
                    break;
                case ERROR:
                    newCell.setCellErrorValue(oldCell.getErrorCellValue());
                    break;
                case FORMULA:
                    newCell.setCellFormula(oldCell.getCellFormula());
                    break;
                case NUMERIC:
                    newCell.setCellValue(oldCell.getNumericCellValue());
                    break;
                case STRING:
                    newCell.setCellValue(oldCell.getRichStringCellValue());
                    break;
                case _NONE:
                    break;
            }
        }

    }
    
    static private Pattern rxquote = Pattern.compile("\"");

    static private String encodeValue(String value) {
        boolean needQuotes = false;
        if ( value.indexOf(',') != -1 || value.indexOf('"') != -1 ||
             value.indexOf('\n') != -1 || value.indexOf('\r') != -1 )
            needQuotes = true;
        Matcher m = rxquote.matcher(value);
        if ( m.find() ) needQuotes = true; value = m.replaceAll("\"\"");
        if ( needQuotes ) return "\"" + value + "\"";
        else return value;
    }
}
