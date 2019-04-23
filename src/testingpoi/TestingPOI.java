/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */

package testingpoi;

import java.io.File;


/**
 *
 * @author miquel.mirat
 */
public class TestingPOI {
    static JavaPoiUtils javaPoiUtils = new JavaPoiUtils();
   
    //javaPoiUtils.readExcelFile(new File("/home/xules/codigoxules/apachepoi/PaisesIdiomasMonedas.xls"));  
    /**
     * @param args the command line arguments
     */
    public static void main(String[] args) {
        javaPoiUtils.readExcelFile(new File("output.xls"));
    }
    
}
