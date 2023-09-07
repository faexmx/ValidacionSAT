/*
 * Click nbfs://nbhost/SystemFileSystem/Templates/Licenses/license-default.txt to change this license
 * Click nbfs://nbhost/SystemFileSystem/Templates/Classes/Class.java to edit this template
 */
package mx.itq.isc.soa.validacionsat;

import java.io.File;
import java.util.Vector;
import javax.xml.parsers.DocumentBuilderFactory;
import javax.xml.parsers.DocumentBuilder;
import javax.xml.validation.Schema;
import javax.xml.validation.SchemaFactory;
import javax.xml.validation.Validator;
import org.w3c.dom.Document;
import org.w3c.dom.Element;
import org.w3c.dom.NodeList;

public class CFDIParser {

    public static void main(String[] args) throws Exception {
        // Leer el archivo XML
        DocumentBuilderFactory factory = DocumentBuilderFactory.newInstance();
        DocumentBuilder builder = factory.newDocumentBuilder();
        File archivo = new File("D:/ITQ/PROYECTOS_NETBEANS/EJEMPLOS_CFDI/63851D4F-3F61-41CE-AF43-8235409953F0.xml");
        //Document document = builder.parse("D:/ITQ/PROYECTOS_NETBEANS/EJEMPLOS_CFDI/63851D4F-3F61-41CE-AF43-8235409953F0.xml");
        Document document = builder.parse(archivo);

        // Validar el archivo XML
//        SchemaFactory schemaFactory = SchemaFactory.newInstance(javax.xml.XMLConstants.W3C_XML_SCHEMA_NS_URI);
//        Schema schema = schemaFactory.newSchema(new File("D:\\ITQ\\PROYECTOS_NETBEANS\\EJEMPLOS_CFDI\\cfdi33.xsd"));
//        Validator validator = schema.newValidator();
        // Obtener la información del archivo XML
        Element root = document.getDocumentElement();

        // RFC Emisor
        String rfcEmisor = root.getElementsByTagName("cfdi:Emisor").item(0).getAttributes().getNamedItem("Rfc").getTextContent();

        // RFC Receptor
        String rfcReceptor = root.getElementsByTagName("cfdi:Receptor").item(0).getAttributes().getNamedItem("Rfc").getTextContent();

        // Total
        String total = root.getAttribute("Total");

        // Total
        String fecha = root.getAttribute("Fecha");

        // UUID
        String uuid = root.getElementsByTagName("tfd:TimbreFiscalDigital").item(0).getAttributes().getNamedItem("UUID").getTextContent();

        // Imprimir la información
        System.out.println("RFC Emisor: " + rfcEmisor);
        System.out.println("RFC Receptor: " + rfcReceptor);
        System.out.println("Total: " + total);
        System.out.println("UUID: " + uuid);
        System.out.println("FECHA EMISION: " + fecha);
    }

    public void cargarCFDIXML(String ubicacionXML) {
//        Vector respuetas = new Vector();
//        DocumentBuilderFactory factory = DocumentBuilderFactory.newInstance();
//        DocumentBuilder builder = factory.newDocumentBuilder();
//        File archivo = new File(ubicacionXML);
//        //Document document = builder.parse("D:/ITQ/PROYECTOS_NETBEANS/EJEMPLOS_CFDI/63851D4F-3F61-41CE-AF43-8235409953F0.xml");
//        Document document = builder.parse(archivo);

    }

    public void validarXML() {

    }

    public void obtenerDatosXML() {

    }
}