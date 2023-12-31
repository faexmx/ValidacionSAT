/*
 * Click nbfs://nbhost/SystemFileSystem/Templates/Licenses/license-default.txt to change this license
 * Click nbfs://nbhost/SystemFileSystem/Templates/Classes/Class.java to edit this template
 */
package mx.itq.isc.soa.validacionsat;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.HashMap;
import java.util.List;
import mx.sat.wsvalidacion.Acuse;
import mx.sat.wsvalidacion.ConsultaCFDIService;
import mx.sat.wsvalidacion.IConsultaCFDIService;
import org.apache.poi.ss.usermodel.Cell;
import static org.apache.poi.ss.usermodel.CellType.BLANK;
import static org.apache.poi.ss.usermodel.CellType.BOOLEAN;
import static org.apache.poi.ss.usermodel.CellType.ERROR;
import static org.apache.poi.ss.usermodel.CellType.FORMULA;
import static org.apache.poi.ss.usermodel.CellType.NUMERIC;
import static org.apache.poi.ss.usermodel.CellType.STRING;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 *
 * @author diazn
 */
public class ValidacionSAT {

    String rutaArchivo = "";
    PantallaInicial pantalla = null;

    public ValidacionSAT(PantallaInicial pantalla) {
        this.pantalla = pantalla;
    }

    public ValidacionSAT() {
    }

    public ValidacionSAT(String rutaArchivo, PantallaInicial pantalla) {
        this.rutaArchivo = rutaArchivo;
        this.pantalla = pantalla;
    }

    public String obtencionEstatus(String rfcEmisor, String rfcReceptor, String total, String uuid) {

        rfcEmisor = rfcEmisor.replaceAll("&", "&amp;");
        rfcReceptor = rfcReceptor.replaceAll("&", "&amp;");
        String respuestaPeticion = "SIN RESPUESTA DEL SAT";
        String cadenaPeticion = "?re=" + rfcEmisor + "&rr=" + rfcReceptor + "&tt=" + total + "&id=" + uuid;
        ValidacionSAT operacionesWSValidacion = new ValidacionSAT(this.pantalla);
        Acuse acuseSAT = operacionesWSValidacion.consulta(cadenaPeticion);
        if (acuseSAT != null) {
            if (acuseSAT.getCodigoEstatus() != null) {
                String codigoEstatus = acuseSAT.getCodigoEstatus().getValue();
                if (codigoEstatus != null && !codigoEstatus.equals("")) {
                    if (codigoEstatus.toUpperCase().equals("N - 601: La expresión impresa proporcionada no es válida.".toUpperCase())) {
                        respuestaPeticion = "SIN RESPUESTA DEL SAT";
                    } else if (codigoEstatus.toUpperCase().equals("N - 601: La expresión impresa proporcionada no es válida.".toUpperCase())) {
                        respuestaPeticion = "NO EXISTE EN SAT";
                    } else if (codigoEstatus.toUpperCase().equals("S - Comprobante obtenido satisfactoriamente.".toUpperCase())) {
                        if (acuseSAT.getEstado() != null) {
                            String estatusSAT = acuseSAT.getEstado().getValue();
                            if (estatusSAT != null && !estatusSAT.equals("")) {
                                if (estatusSAT.toUpperCase().equals("VIGENTE")) {
                                    respuestaPeticion = "VIGENTE";
                                } else if (estatusSAT.toUpperCase().equals("CANCELADO")) {
                                    respuestaPeticion = "CANCELADO";
                                }
                            } else {
                                respuestaPeticion = "NO EXISTE EN SAT";
                            }
                        } else {
                            respuestaPeticion = "NO EXISTE EN SAT";
                        }
                    }
                } else {
                    respuestaPeticion = "SIN RESPUESTA DEL SAT";
                }
            } else {
                respuestaPeticion = "SIN RESPUESTA DEL SAT";
            }
        } else {
            respuestaPeticion = "SIN RESPUESTA DEL SAT";
        }
        return respuestaPeticion;
    }

    public Acuse consulta(java.lang.String expresionImpresa) {
        try {
            ConsultaCFDIService service = new ConsultaCFDIService();
            IConsultaCFDIService port = service.getBasicHttpBindingIConsultaCFDIService();
            return port.consulta(expresionImpresa);
        } catch (Exception ex) {
            ex.printStackTrace();
            return null;
        }
    }

    public File cargaArchivoListadoCFDI(String ubicacionArchivo) {
        File archivo = null;

        return archivo;
    }

    public String datosCFDI(String rfcEmisor, String rfcReceptor, String total, String uuid) {

        rfcEmisor = rfcEmisor.replaceAll("&", "&amp;");
        rfcReceptor = rfcReceptor.replaceAll("&", "&amp;");
        String respuestaPeticion = "SIN RESPUESTA DEL SAT";
        String cadenaPeticion = "?re=" + rfcEmisor + "&rr=" + rfcReceptor + "&tt=" + total + "&id=" + uuid;
        ValidacionSAT operacionesWSValidacion = new ValidacionSAT(this.pantalla);
        Acuse acuseSAT = operacionesWSValidacion.consulta(cadenaPeticion);
        if (acuseSAT != null) {
            if (acuseSAT.getCodigoEstatus() != null) {
                String codigoEstatus = acuseSAT.getCodigoEstatus().getValue();
                if (codigoEstatus != null && !codigoEstatus.equals("")) {
                    if (codigoEstatus.toUpperCase().equals("N - 601: La expresión impresa proporcionada no es válida.".toUpperCase())) {
                        respuestaPeticion = "SIN RESPUESTA DEL SAT";
                    } else if (codigoEstatus.toUpperCase().equals("N - 601: La expresión impresa proporcionada no es válida.".toUpperCase())) {
                        respuestaPeticion = "NO EXISTE EN SAT";
                    } else if (codigoEstatus.toUpperCase().equals("S - Comprobante obtenido satisfactoriamente.".toUpperCase())) {
                        if (acuseSAT.getEstado() != null) {
                            String estatusSAT = acuseSAT.getEstado().getValue();
                            if (estatusSAT != null && !estatusSAT.equals("")) {
                                if (estatusSAT.toUpperCase().equals("VIGENTE")) {
                                    respuestaPeticion = "VIGENTE";
                                } else if (estatusSAT.toUpperCase().equals("CANCELADO")) {
                                    respuestaPeticion = "CANCELADO";
                                }
                            } else {
                                respuestaPeticion = "NO EXISTE EN SAT";
                            }
                        } else {
                            respuestaPeticion = "NO EXISTE EN SAT";
                        }
                    }
                } else {
                    respuestaPeticion = "SIN RESPUESTA DEL SAT";
                }
            } else {
                respuestaPeticion = "SIN RESPUESTA DEL SAT";
            }
        } else {
            respuestaPeticion = "SIN RESPUESTA DEL SAT";
        }
        return respuestaPeticion;
    }

    public String validarExcel(String rutaExcel) {
        String resultadoProcesamiento = "CORRECTO";
        try {
            // Ruta al archivo Excel
            this.pantalla.escribirConsola("*** INICIO VALIDACION");
            String excelFilePath = rutaExcel;
            // Cargar el archivo Excel
            FileInputStream inputStream = new FileInputStream(excelFilePath);
            Workbook workbook = new XSSFWorkbook(inputStream);
            // Obtener la hoja de trabajo (worksheet) que deseas leer
            Sheet sheet = workbook.getSheetAt(0); // 0 representa la primera hoja
            // Iterar a través de las filas y columnas para obtener los datos
            int contadorRenglon = 0;//contador para indicar numero de renglon
            int contadorCelda = 1;//contador para indicar numero de celda
            String rfcEmisor = "";
            String rfcReceptor = "";
            String total = "";
            String uuid = "";
            String fecha = "";
            String serie = "";
            String folio = "";

            for (Row row : sheet) {
                contadorRenglon++;
                if (contadorRenglon == 1) {
                    continue;
                }
                //SE VA A RECORRER LAS CELDAS DE LA FILA PARA ARMAR CADENA DE PETICION
                rfcEmisor = "";
                rfcReceptor = "";
                total = "";
                uuid = "";
                HashMap mapDatosSAT = new HashMap();
                contadorCelda = 1;
                mapDatosSAT.put("RESULTADO", "SIN RESPUESTA DEL SAT");
                for (Cell cell : row) {
                    // Obtener el valor de la celda como una cadena
                    String cellValue = cell.toString();
                    System.out.print(cellValue + "\t");
                    if (contadorCelda == 1) {
                        rfcEmisor = cellValue;
                    }
                    if (contadorCelda == 2) {
                        rfcReceptor = cellValue;
                    }
                    if (contadorCelda == 3) {
                        total = cellValue;
                    }
                    if (contadorCelda == 4) {
                        uuid = cellValue;
                    }
                    if (contadorCelda == 5) {
                        fecha = cellValue;
                    }
                    if (contadorCelda == 6) {
                        serie = cellValue;
                    }
                    if (contadorCelda == 7) {
                        folio = cellValue;
                    }

                    if (contadorCelda == 8) {

                        mapDatosSAT = obtencionEstatusSAT(rfcEmisor, rfcReceptor, total, uuid);
                        this.pantalla.escribirConsola((contadorRenglon - 1)
                                + " -  RFC EMISOR = " + rfcEmisor
                                + ", RFC RECEPTOR = " + rfcReceptor
                                + ", TOTAL = " + total
                                + ", UUID = " + uuid
                                + ", ESTATUS PROCESO = " + mapDatosSAT.get("RESULTADO").toString()
                        );
                    }

                    if (contadorCelda == 8) {
                        cell.setCellValue(mapDatosSAT.get("ESTATUSPETICION").toString());
                    }

                    if (mapDatosSAT.get("RESULTADO").toString().equals("S - Comprobante obtenido satisfactoriamente.")) {
                        if (contadorCelda == 9) {
                            cell.setCellValue(mapDatosSAT.get("ES1TATUSCFDI").toString());
                        }
                        if (contadorCelda == 10) {
                            cell.setCellValue(mapDatosSAT.get("ESCANCELABLE").toString());
                        }
                        if (contadorCelda == 11) {
                            cell.setCellValue(mapDatosSAT.get("ESTATUSCANCELACION").toString());
                        }
                        if (contadorCelda == 12) {
                            cell.setCellValue(mapDatosSAT.get("VALIDACIONEFOS").toString());
                        }
                    }
                    contadorCelda++;
                }

                System.out.println(); // Nueva línea para cada fila
            }
            // Cerrar el flujo de entrada
            inputStream.close();
            FileOutputStream outputStream = new FileOutputStream(rutaExcel);
            workbook.write(outputStream);
            outputStream.close();
            // Cerrar el libro de Excel
            workbook.close();
            this.pantalla.escribirConsola("*** NUMERO DE REGISTROS PROCESADOS = " + (contadorRenglon - 1));

            System.out.println("Archivo Excel creado con éxito.");
            // Puedes almacenar los datos en una variable de acuerdo a tus necesidades
        } catch (IOException e) {
            e.printStackTrace();
        } finally {
            return resultadoProcesamiento;
        }
    }

    public String validacionYCargaArchivosXML(File[] ficherosXML, String carpeta) {
        List<HashMap> listaMapaDatosCFDIXML = new ArrayList();

        for (File fichero : ficherosXML) {
            System.out.println("Archivo seleccionado: " + fichero.getAbsolutePath());
            CFDIParserXML parseadorCFDI = new CFDIParserXML();
            HashMap mapaCFDI = parseadorCFDI.obtenerDatosXML(fichero.getAbsolutePath());
            listaMapaDatosCFDIXML.add(mapaCFDI);
        }

        for (HashMap mapaCFDIXML : listaMapaDatosCFDIXML) {

            HashMap mapaRespuestaSAT = obtencionEstatusSAT(mapaCFDIXML.get("RFCEMISOR").toString(),
                    mapaCFDIXML.get("RFCRECEPTOR").toString(),
                    mapaCFDIXML.get("TOTAL").toString(),
                    mapaCFDIXML.get("UUID").toString());

            // Juntar los dos mapas
            mapaCFDIXML.putAll(mapaRespuestaSAT);
        }

        generaExcelAPartirDeCFIXML(listaMapaDatosCFDIXML, carpeta);
        return "PPROCESO TERMINADO";
    }

    public void generaExcelAPartirDeCFIXML(List<HashMap> listaMapaDatosCFDIXML, String carpeta) {

        try {
            // Crear un nuevo Excel
            Workbook workbook = new XSSFWorkbook();
            // Crear una nueva hoja
            Sheet sheet = workbook.createSheet("LISTADO CFDI");
            // Crear encabezados
            Row headerRow = sheet.createRow(0);
            headerRow.createCell(0).setCellValue("RFC EMISOR");
            headerRow.createCell(1).setCellValue("RFC EMISOR");
            headerRow.createCell(2).setCellValue("TOTAL");
            headerRow.createCell(3).setCellValue("UUID");
            headerRow.createCell(4).setCellValue("FECHA");
            headerRow.createCell(5).setCellValue("SERIE");
            headerRow.createCell(6).setCellValue("FOLIO");
            headerRow.createCell(7).setCellValue("RESPUESTA SAT PROCESO");
            headerRow.createCell(8).setCellValue("ESTATUS CFDI ");
            headerRow.createCell(9).setCellValue("ES CANCELABLE");
            headerRow.createCell(10).setCellValue("ESTATUS CANCELACION");
            headerRow.createCell(11).setCellValue("EFOS");

            // Iterar sobre la lista de HashMaps
            int contador = 1;
            for (HashMap mapaCFDIXML : listaMapaDatosCFDIXML) {
                // Iterar sobre el HashMap actual
                // Crear una nueva fila
                Row row = sheet.createRow(sheet.getLastRowNum() + 1);
                row.createCell(0).setCellValue(mapaCFDIXML.get("RFCEMISOR").toString());
                row.createCell(1).setCellValue(mapaCFDIXML.get("RFCRECEPTOR").toString());
                row.createCell(2).setCellValue(mapaCFDIXML.get("TOTAL").toString());
                row.createCell(3).setCellValue(mapaCFDIXML.get("UUID").toString());
                row.createCell(4).setCellValue(mapaCFDIXML.get("FECHA").toString());
                row.createCell(5).setCellValue(mapaCFDIXML.get("SERIE").toString());
                row.createCell(6).setCellValue(mapaCFDIXML.get("FOLIO").toString());
                if (mapaCFDIXML.get("ESTATUSPETICION") != null) {
                    row.createCell(7).setCellValue(mapaCFDIXML.get("ESTATUSPETICION").toString());
                } else {
                    row.createCell(7).setCellValue("");
                }
                if (mapaCFDIXML.get("ES1TATUSCFDI") != null) {
                    row.createCell(8).setCellValue(mapaCFDIXML.get("ES1TATUSCFDI").toString());
                } else {
                    row.createCell(8).setCellValue("");
                }
                if (mapaCFDIXML.get("ESCANCELABLE") != null) {
                    row.createCell(9).setCellValue(mapaCFDIXML.get("ESCANCELABLE").toString());
                } else {
                    row.createCell(9).setCellValue("");
                }
                if (mapaCFDIXML.get("ESTATUSCANCELACION") != null) {
                    row.createCell(10).setCellValue(mapaCFDIXML.get("ESTATUSCANCELACION").toString());
                } else {
                    row.createCell(10).setCellValue("");
                }

                if (mapaCFDIXML.get("VALIDACIONEFOS") != null) {
                    row.createCell(11).setCellValue(mapaCFDIXML.get("VALIDACIONEFOS").toString());
                } else {
                    row.createCell(11).setCellValue("");
                }

                String cadenaExcel = contador + ".- ";
                if (row != null) { // Verifica si la fila no es nula
                    for (Cell cell : row) {
                        switch (cell.getCellType()) {
                            case STRING:
                                System.out.print(cell.getStringCellValue() + " ");
                                cadenaExcel = cadenaExcel + cell.getStringCellValue() + " ";
                                break;
                            case NUMERIC:
                                System.out.print(cell.getNumericCellValue() + " ");
                                cadenaExcel = cadenaExcel + cell.getNumericCellValue() + " ";
                                break;
                            case BOOLEAN:
                                System.out.print(cell.getBooleanCellValue() + " ");
                                cadenaExcel = cadenaExcel + cell.getBooleanCellValue() + " ";
                                break;
                            case BLANK:
                                System.out.print("[BLANK]\t");
                                cadenaExcel = cadenaExcel + "[BLANK] ";
                                break;
                            case ERROR:
                                System.out.print("[ERROR]\t");
                                cadenaExcel = cadenaExcel + "[ERROR] ";
                                break;
                            case FORMULA:
                                System.out.print(cell.getCellFormula() + " ");
                                cadenaExcel = cadenaExcel + cell.getCellFormula() + " ";
                                break;
                            default:
                                System.out.print("[UNKNOWN] ");
                                cadenaExcel = cadenaExcel + cell.getCellFormula() + "\t";
                        }
                    }
                    System.out.println();
                }
                cadenaExcel = cadenaExcel + "\n";
                this.pantalla.escribirConsola(cadenaExcel);
                contador++;

            }
            Date fecha = new Date();
            SimpleDateFormat sdf = new SimpleDateFormat("yyyyMMdd_HHmmss");
            String fechaFormateada = sdf.format(fecha);
            FileOutputStream outputStream = new FileOutputStream(carpeta + "/LISTADO_CLASE_CFDI_" + fechaFormateada + ".xlsx");
            workbook.write(outputStream);
            outputStream.close();
            workbook.close();
        } catch (Exception ex) {
            ex.printStackTrace();
        }
    }

    public HashMap obtencionEstatusSAT(String rfcEmisor, String rfcReceptor, String total, String uuid) {
        HashMap mapDatosSAT = new HashMap();
        rfcEmisor = rfcEmisor.replaceAll("&", "&amp;");
        rfcReceptor = rfcReceptor.replaceAll("&", "&amp;");
        String respuestaPeticion = "SIN RESPUESTA DEL SAT";
        String cadenaPeticion = "?re=" + rfcEmisor + "&rr=" + rfcReceptor + "&tt=" + total + "&id=" + uuid;
        ValidacionSAT operacionesWSValidacion = new ValidacionSAT(this.pantalla);
        Acuse acuseSAT = operacionesWSValidacion.consulta(cadenaPeticion);
        if (acuseSAT != null) {
            if (acuseSAT.getCodigoEstatus() != null) {
                String codigoEstatus = acuseSAT.getCodigoEstatus().getValue();
                if (codigoEstatus != null && !codigoEstatus.equals("")) {
                    mapDatosSAT.put("ESTATUSPETICION", codigoEstatus);

                    if (codigoEstatus.toUpperCase().equals("N - 601: La expresión impresa proporcionada no es válida.".toUpperCase())) {
                        respuestaPeticion = "N - 601: La expresión impresa proporcionada no es válida.";
                    } else if (codigoEstatus.toUpperCase().equals("N - 601: La expresión impresa proporcionada no es válida.".toUpperCase())) {
                        respuestaPeticion = "N - 602: Comprobante no encontrado. ";
                    } else if (codigoEstatus.toUpperCase().equals("S - Comprobante obtenido satisfactoriamente.".toUpperCase())) {
                        if (acuseSAT.getEstado() != null) {
                            mapDatosSAT.put("ES1TATUSCFDI", acuseSAT.getEstado().getValue());
                            mapDatosSAT.put("ESCANCELABLE", acuseSAT.getEsCancelable().getValue());
                            mapDatosSAT.put("ESTATUSCANCELACION", acuseSAT.getEstatusCancelacion().getValue());
                            mapDatosSAT.put("VALIDACIONEFOS", acuseSAT.getValidacionEFOS().getValue());
                            respuestaPeticion = "S - Comprobante obtenido satisfactoriamente.";
                        } else {
                            respuestaPeticion = "NO EXISTE EN SAT";
                        }
                    }
                } else {
                    respuestaPeticion = "SIN RESPUESTA DEL SAT";
                }
            } else {
                respuestaPeticion = "SIN RESPUESTA DEL SAT";
            }
        } else {
            respuestaPeticion = "SIN RESPUESTA DEL SAT";
        }

        mapDatosSAT.put("RESULTADO", respuestaPeticion);
        return mapDatosSAT;
    }

    public Workbook generaExcelAPartirDeCFDIXMLWEB(List<HashMap> listaMapaDatosCFDIXML) {

        Workbook workbook = null;
        try {
            // Crear un nuevo Excel
            workbook = new XSSFWorkbook();
            // Crear una nueva hoja
            Sheet sheet = workbook.createSheet("LISTADO CFDI");
            // Crear encabezados
            Row headerRow = sheet.createRow(0);
            headerRow.createCell(0).setCellValue("RFC EMISOR");
            headerRow.createCell(1).setCellValue("RFC EMISOR");
            headerRow.createCell(2).setCellValue("TOTAL");
            headerRow.createCell(3).setCellValue("UUID");
            headerRow.createCell(4).setCellValue("FECHA");
            headerRow.createCell(5).setCellValue("SERIE");
            headerRow.createCell(6).setCellValue("FOLIO");
            headerRow.createCell(7).setCellValue("RESPUESTA SAT PROCESO");
            headerRow.createCell(8).setCellValue("ESTATUS CFDI ");
            headerRow.createCell(9).setCellValue("ES CANCELABLE");
            headerRow.createCell(10).setCellValue("ESTATUS CANCELACION");
            headerRow.createCell(11).setCellValue("EFOS");

            // Iterar sobre la lista de HashMaps
            int contador = 1;
            for (HashMap mapaCFDIXML : listaMapaDatosCFDIXML) {
                // Iterar sobre el HashMap actual
                // Crear una nueva fila
                Row row = sheet.createRow(sheet.getLastRowNum() + 1);
                row.createCell(0).setCellValue(mapaCFDIXML.get("RFCEMISOR").toString());
                row.createCell(1).setCellValue(mapaCFDIXML.get("RFCRECEPTOR").toString());
                row.createCell(2).setCellValue(mapaCFDIXML.get("TOTAL").toString());
                row.createCell(3).setCellValue(mapaCFDIXML.get("UUID").toString());
                row.createCell(4).setCellValue(mapaCFDIXML.get("FECHA").toString());
                row.createCell(5).setCellValue(mapaCFDIXML.get("SERIE").toString());
                row.createCell(6).setCellValue(mapaCFDIXML.get("FOLIO").toString());
                if (mapaCFDIXML.get("ESTATUSPETICION") != null) {
                    row.createCell(7).setCellValue(mapaCFDIXML.get("ESTATUSPETICION").toString());
                } else {
                    row.createCell(7).setCellValue("");
                }
                if (mapaCFDIXML.get("ES1TATUSCFDI") != null) {
                    row.createCell(8).setCellValue(mapaCFDIXML.get("ES1TATUSCFDI").toString());
                } else {
                    row.createCell(8).setCellValue("");
                }
                if (mapaCFDIXML.get("ESCANCELABLE") != null) {
                    row.createCell(9).setCellValue(mapaCFDIXML.get("ESCANCELABLE").toString());
                } else {
                    row.createCell(9).setCellValue("");
                }
                if (mapaCFDIXML.get("ESTATUSCANCELACION") != null) {
                    row.createCell(10).setCellValue(mapaCFDIXML.get("ESTATUSCANCELACION").toString());
                } else {
                    row.createCell(10).setCellValue("");
                }

                if (mapaCFDIXML.get("VALIDACIONEFOS") != null) {
                    row.createCell(11).setCellValue(mapaCFDIXML.get("VALIDACIONEFOS").toString());
                } else {
                    row.createCell(11).setCellValue("");
                }

                String cadenaExcel = contador + ".- ";
                if (row != null) { // Verifica si la fila no es nula
                    for (Cell cell : row) {
                        switch (cell.getCellType()) {
                            case STRING:
                                System.out.print(cell.getStringCellValue() + " ");
                                cadenaExcel = cadenaExcel + cell.getStringCellValue() + " ";
                                break;
                            case NUMERIC:
                                System.out.print(cell.getNumericCellValue() + " ");
                                cadenaExcel = cadenaExcel + cell.getNumericCellValue() + " ";
                                break;
                            case BOOLEAN:
                                System.out.print(cell.getBooleanCellValue() + " ");
                                cadenaExcel = cadenaExcel + cell.getBooleanCellValue() + " ";
                                break;
                            case BLANK:
                                System.out.print("[BLANK]\t");
                                cadenaExcel = cadenaExcel + "[BLANK] ";
                                break;
                            case ERROR:
                                System.out.print("[ERROR]\t");
                                cadenaExcel = cadenaExcel + "[ERROR] ";
                                break;
                            case FORMULA:
                                System.out.print(cell.getCellFormula() + " ");
                                cadenaExcel = cadenaExcel + cell.getCellFormula() + " ";
                                break;
                            default:
                                System.out.print("[UNKNOWN] ");
                                cadenaExcel = cadenaExcel + cell.getCellFormula() + "\t";
                        }
                    }
                    System.out.println();
                }
                cadenaExcel = cadenaExcel + "\n";
                System.out.println(cadenaExcel);
                contador++;

            }
            Date fecha = new Date();
            SimpleDateFormat sdf = new SimpleDateFormat("yyyyMMdd_HHmmss");
            String fechaFormateada = sdf.format(fecha);
            //FileOutputStream outputStream = new FileOutputStream(carpeta + "/LISTADO_CLASE_CFDI_" + fechaFormateada + ".xlsx");
            //workbook.write(outputStream);
            //outputStream.close();
            //workbook.close();
        } catch (Exception ex) {
            ex.printStackTrace();
        }
        return workbook;
    }

}
