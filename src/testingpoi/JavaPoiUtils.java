/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */

package testingpoi;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.InputStream;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook; 
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 *
 * @author miquel.mirat
 */
public class JavaPoiUtils {
   
    public void readExcelFile(File excelFile){
        InputStream excelStream = null;
        try {
            excelStream = new FileInputStream(excelFile);
            // High level representation of a workbook.
            // Representación del más alto nivel de la hoja excel.
            //HSSFWorkbook hssfWorkbook = new HSSFWorkbook(excelStream);
            XSSFWorkbook xssfWorkbook = new XSSFWorkbook(excelStream);
            // We chose the sheet is passed as parameter. 
            // Elegimos la hoja que se pasa por parámetro.
            XSSFSheet sheet = xssfWorkbook.getSheetAt(0);
            // An object that allows us to read a row of the excel sheet, and extract from it the cell contents.
            // Objeto que nos permite leer un fila de la hoja excel, y de aquí extraer el contenido de las celdas.

            XSSFRow row;
            // Initialize the object to read the value of the cell 
            // Inicializo el objeto que leerá el valor de la celda
            XSSFCell cell;                        
            // I get the number of rows occupied on the sheet
            // Obtengo el número de filas ocupadas en la hoja
            int rows = sheet.getLastRowNum();
            // I get the number of columns occupied on the sheet
            // Obtengo el número de columnas ocupadas en la hoja
            int cols = 0;            
            // A string used to store the reading cell
            // Cadena que usamos para almacenar la lectura de la celda
            String cellValue = "";  
            // For this example we'll loop through the rows getting the data we want
            // Para este ejemplo vamos a recorrer las filas obteniendo los datos que queremos            
            for (int r = 0; r < rows; r++) {
                row = sheet.getRow(r);
                if (row == null){
                    break;
                }else{
                    System.out.print("Row: " + r + " -> ");
                    for (int c = 0; c < (cols = row.getLastCellNum()); c++) {
                        /* 
                            We have those cell types (tenemos estos tipos de celda): 
                                CELL_TYPE_BLANK, CELL_TYPE_NUMERIC, CELL_TYPE_BLANK, CELL_TYPE_FORMULA, CELL_TYPE_BOOLEAN, CELL_TYPE_ERROR
                        */
                        if (row.getCell(c) != null) {
                            //System.out.println(row.getCell(c).getCellType());
                            if (row.getCell(c).getCellType() == CellType.STRING) {
                                cellValue = row.getCell(c).getStringCellValue();
                            } else if (row.getCell(c).getCellType() == CellType.NUMERIC) {
                                cellValue = String.valueOf(row.getCell(c).getNumericCellValue());
                            }

                        } else {
                            cellValue = "";
                        }

//                                (row.getCell(c).getCellType() == Cell.CELL_TYPE_STRING)?row.getCell(c).getStringCellValue():
//                                (row.getCell(c).getCellType() == Cell.CELL_TYPE_NUMERIC)?"" + row.getCell(c).getNumericCellValue():
//                                (row.getCell(c).getCellType() == Cell.CELL_TYPE_BOOLEAN)?"" + row.getCell(c).getBooleanCellValue():
//                                (row.getCell(c).getCellType() == Cell.CELL_TYPE_BLANK)?"BLANK":
//                                (row.getCell(c).getCellType() == Cell.CELL_TYPE_FORMULA)?"FORMULA":
//                                (row.getCell(c).getCellType() == Cell.CELL_TYPE_ERROR)?"ERROR":"";                       
                        System.out.print("[Column " + c + ": " + cellValue + "] ");
                    }
                }
            }            
        } catch (FileNotFoundException fileNotFoundException) {
            System.out.println("The file not exists (No se encontró el fichero): " + fileNotFoundException);
        } catch (IOException ex) {
            System.out.println("Error in file procesing (Error al procesar el fichero): " + ex);
        } finally {
            try {
                excelStream.close();
            } catch (IOException ex) {
                System.out.println("Error in file processing after close it (Error al procesar el fichero después de cerrarlo): " + ex);
            }
        }
    }
    /**     
     * Main method for the tests for the methods of the class <strong>Java
     * read excel</strong> and <strong>Java create excel</strong> 
     * with <a href="https://poi.apache.org/">Apache POI</a>. 
     * <br />
     * Método main para las pruebas para los método de la clase,
     * pruebas de <strong>Java leer excel</strong> y  <strong>Java crear excel</strong>
     * con <a href="https://poi.apache.org/">Apache POI</a>.     
     * @param args 
     */
//    public static void main(String[] args){
//        JavaPoiUtils javaPoiUtils = new JavaPoiUtils();
//        javaPoiUtils.readExcelFile(new File("/home/xules/codigoxules/apachepoi/PaisesIdiomasMonedas.xls"));        
//    }    
    
    
    
    
    ///codigo de ejemplo!
    
//    XSSFWorkbook wb = new XSSFWorkbook();
//    XSSFSheet sheet = wb.createSheet();
//    XSSFRow row = sheet.createRow(0);
//    XSSFCell cell = row.createCell( 0);
//    cell.setCellValue("custom XSSF colors");
//
//    XSSFCellStyle style1 = wb.createCellStyle();
//    style1.setFillForegroundColor(new XSSFColor(new java.awt.Color(128, 0, 128), new DefaultIndexedColorMap()));
//    style1.setFillPattern(FillPatternType.SOLID_FOREGROUND);
//                    
//Reading and Rewriting Workbooks
}
