/*
 * Click nbfs://nbhost/SystemFileSystem/Templates/Licenses/license-default.txt to change this license
 */
package mx.itq.isc.soa.validacionsat;

import java.awt.Dimension;
import java.awt.Toolkit;
import mx.sat.wsvalidacion.Acuse;
import mx.sat.wsvalidacion.ConsultaCFDIService;
import mx.sat.wsvalidacion.IConsultaCFDIService;

/**
 *
 * @author diazn
 */
public class CONEXION_SATF {

    public static void main(String[] args) {
        System.out.println("HOLA SOA!");
        PantallaInicial pantallaIni = new PantallaInicial();
        pantallaIni.setSize(780,600);
        Dimension pantalla = Toolkit.getDefaultToolkit().getScreenSize();
        Dimension frame = pantallaIni.getSize();
        pantallaIni.setLocation((pantalla.width / 2 - (frame.width / 2)), pantalla.height / 2 - (frame.height / 2));
        pantallaIni.setVisible(true);        
    }   

}
