/*
 * Click nbfs://nbhost/SystemFileSystem/Templates/Licenses/license-default.txt to change this license
 * Click nbfs://nbhost/SystemFileSystem/Templates/Classes/Class.java to edit this template
 */
package mx.itq.isc.soa.validacionsat;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.IOException;

public class ExcelLoader {
    
    
    public static void main(String[] args) {
        try {
            // Ruta al archivo Excel
            String excelFilePath = "D:\\ITQ\\PROYECTOS_NETBEANS\\LISTADOS_ENTRADA\\LAYOUT_ENTRADA.xlsx";
            
            // Cargar el archivo Excel
            FileInputStream inputStream = new FileInputStream(excelFilePath);
            Workbook workbook = new XSSFWorkbook(inputStream);

            // Obtener la hoja de trabajo (worksheet) que deseas leer
            Sheet sheet = workbook.getSheetAt(0); // 0 representa la primera hoja

            // Iterar a través de las filas y columnas para obtener los datos
            for (Row row : sheet) {
                for (Cell cell : row) {
                    // Obtener el valor de la celda como una cadena
                    String cellValue = cell.toString();
                    System.out.print(cellValue + "\t");
                }
                System.out.println(); // Nueva línea para cada fila
            }

            // Cerrar el flujo de entrada
            inputStream.close();

            // Puedes almacenar los datos en una variable de acuerdo a tus necesidades

        } catch (IOException e) {
            e.printStackTrace();
        }
    }
    
        public String validarExcel(String rutaExcel) {
        String resultadoProcesamiento = "CORRECTo";
        try {
            // Ruta al archivo Excel
            String excelFilePath = rutaExcel;            
            // Cargar el archivo Excel
            FileInputStream inputStream = new FileInputStream(excelFilePath);
            Workbook workbook = new XSSFWorkbook(inputStream);
            // Obtener la hoja de trabajo (worksheet) que deseas leer
            Sheet sheet = workbook.getSheetAt(0); // 0 representa la primera hoja
            // Iterar a través de las filas y columnas para obtener los datos
            for (Row row : sheet) {
                for (Cell cell : row) {
                    // Obtener el valor de la celda como una cadena
                    String cellValue = cell.toString();
                    System.out.print(cellValue + "\t");
                }
                System.out.println(); // Nueva línea para cada fila
            }
            // Cerrar el flujo de entrada
            inputStream.close();
            // Puedes almacenar los datos en una variable de acuerdo a tus necesidades
        } catch (IOException e) {
            e.printStackTrace();
        }finally{
            return resultadoProcesamiento;
        }        
    }
}
