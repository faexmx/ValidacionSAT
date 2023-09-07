/*
 * Click nbfs://nbhost/SystemFileSystem/Templates/Licenses/license-default.txt to change this license
 * Click nbfs://nbhost/SystemFileSystem/Templates/Classes/Class.java to edit this template
 */
package mx.itq.isc.soa.validacionsat;

/**
 *
 * @author diazn
 */
public class Convertidor {

    public static void main(String[] args) {
        String input = "E|Y_461576|Y|461576|2023-08-30T07:30:15|01||217,840000|0,000000|||MXN|217,840000|PUE|77710|01|||||616|||S01|||I|HPP980205ER1_MATRIZ|Carretera chetumal - Pto Juarez Km 309 solidaridad Quintana roo||||Playa del carmen||||MEXICO|77710|XAXX010101000|Ventas Publico General|Carretera chetumal - Pto Juarez Km 309 solidaridad Quintana roo||||Playa del carmen||||MEXICO|77710||0,000000|IBEROSTAR_ALOJADOS_40||||||||||Iberostar Selection Paraiso Maya|23934/2023|23/08/2023||DELTA VACATIONS|||DELTA VACATIONS|Iberostar Selection Paraiso Maya|1,000000||11237|RL18||IGRON,IRINA|0,000000||217,840000||||||||||0,000000||||||23/08/2023|||||||||0%|0,000000|0,000000|217,840000||||||||||||||^C|Y_461576|1|7|SERVICIO|ISH|SANEAMIENTO MPIO SOLIDARIDAD 07NTS|31,120000|217,840000|90111501|E48|0|||01|||||||IGRON,IRINA|JS|23934/2023||||23/08/2023|30/08/2023|08/05/2023|23/08/2023|217,840000|00003539416P|0,000000|6639||||^CI|Y_461576|1|T|217,840000|002|Tasa|0,000000||0,000000|||||^IG|Y_461576|1|T|002|Tasa|0,000000|0,000000|0,000000|||||^End";

        // Utilizamos una expresión regular para encontrar números con comas
        String regex = "(\\d+),(\\d{3})";
        String replaced = input.replaceAll(regex, "$1.$2");

        System.out.println(replaced);
    }

}


