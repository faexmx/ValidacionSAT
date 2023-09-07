/*
 * Click nbfs://nbhost/SystemFileSystem/Templates/Licenses/license-default.txt to change this license
 * Click nbfs://nbhost/SystemFileSystem/Templates/Classes/Class.java to edit this template
 */
package mx.itq.isc.soa.validacionsat;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.HashMap;
import javax.net.ssl.HttpsURLConnection;
import javax.net.ssl.SSLContext;
import javax.net.ssl.TrustManager;
import javax.net.ssl.X509TrustManager;
import mx.sat.wsvalidacion.Acuse;
import mx.sat.wsvalidacion.ConsultaCFDIService;
import mx.sat.wsvalidacion.IConsultaCFDIService;
import org.apache.poi.ss.usermodel.Cell;
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
        ValidacionSAT operacionesWSValidacion = new ValidacionSAT();
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
        ValidacionSAT operacionesWSValidacion = new ValidacionSAT();
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

                        mapDatosSAT = obtencionEstatusSAT(rfcEmisor, rfcReceptor, total, uuid);
                        this.pantalla.escribirConsola((contadorRenglon-1)
                                + " -  RFC EMISOR = " + rfcEmisor
                                + ", RFC RECEPTOR = " + rfcReceptor
                                + ", TOTAL = " + total
                                + ", UUID = " + uuid
                                + ", ESTATUS PROCESO = " + mapDatosSAT.get("RESULTADO").toString()
                        );
                    }

                    if (contadorCelda == 5) {
                        cell.setCellValue(mapDatosSAT.get("ESTATUSPETICION").toString());
                    }

                    if (mapDatosSAT.get("RESULTADO").toString().equals("S - Comprobante obtenido satisfactoriamente.")) {
                        if (contadorCelda == 6) {
                            cell.setCellValue(mapDatosSAT.get("ES1TATUSCFDI").toString());
                        }
                        if (contadorCelda == 7) {
                            cell.setCellValue(mapDatosSAT.get("ESCANCELABLE").toString());
                        }
                        if (contadorCelda == 8) {
                            cell.setCellValue(mapDatosSAT.get("ESTATUSCANCELACION").toString());
                        }
                        if (contadorCelda == 9) {
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

    public HashMap obtencionEstatusSAT(String rfcEmisor, String rfcReceptor, String total, String uuid) {
        HashMap mapDatosSAT = new HashMap();
        rfcEmisor = rfcEmisor.replaceAll("&", "&amp;");
        rfcReceptor = rfcReceptor.replaceAll("&", "&amp;");
        String respuestaPeticion = "SIN RESPUESTA DEL SAT";
        String cadenaPeticion = "?re=" + rfcEmisor + "&rr=" + rfcReceptor + "&tt=" + total + "&id=" + uuid;
        ValidacionSAT operacionesWSValidacion = new ValidacionSAT();
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

}