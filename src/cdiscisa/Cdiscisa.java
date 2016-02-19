/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package cdiscisa;

//import static com.sun.org.apache.bcel.internal.util.SecuritySupport.getResourceAsStream;
import java.awt.Color;
import java.awt.image.BufferedImage;
import java.io.BufferedInputStream;


import java.io.File;
import java.io.FileInputStream;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
//import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.InputStream;


import java.text.DateFormat;
import java.text.Format;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Calendar;


import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.pdmodel.PDPage;
import org.apache.pdfbox.pdmodel.PDPageContentStream;
import org.apache.pdfbox.pdmodel.font.PDFont;
import org.apache.pdfbox.pdmodel.font.PDType1Font;
import org.apache.pdfbox.pdmodel.font.PDType0Font;
//import org.apache.pdfbox.pdmodel.font.encoding;

import org.apache.pdfbox.util.Matrix;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

import java.util.Date;
import java.util.HashMap;
import java.util.Iterator;
import java.util.ListIterator;
import java.util.Locale;
import java.util.Map;
import java.util.SortedSet;
import java.util.TreeSet;
import javax.imageio.ImageIO;
import javax.swing.JOptionPane;
import org.apache.pdfbox.cos.COSDictionary;
import org.apache.pdfbox.io.MemoryUsageSetting;
import org.apache.pdfbox.multipdf.PDFMergerUtility;
import org.apache.pdfbox.pdmodel.graphics.image.JPEGFactory;
import org.apache.pdfbox.pdmodel.graphics.image.LosslessFactory;
import org.apache.pdfbox.pdmodel.graphics.image.PDImageXObject;
import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.util.IOUtils;


/**
 *
 * @author Ernesto Armendáriz Bernal.
 */

class StreamUtil {

    public static final String PREFIX = "temp";
    public static final String SUFFIX = ".png";

    public static File stream2file (InputStream in) throws IOException {
        final File tempFile = File.createTempFile(PREFIX, SUFFIX);
        tempFile.deleteOnExit();
        try (FileOutputStream out = new FileOutputStream(tempFile)) {
            IOUtils.copy(in, out);
        }
        return tempFile;
    }

}

class Directorio {
String determinante,formato,unidad,estado,municipio,direccion;
String razon_social, sucursal, nombre_comercial,dirección_fiscal, RFC,telefono, nombre_contacto,mail_envio_constancias,	mail_cco,notas;

public Directorio(){
    this.determinante = ""; // o número de cliente
    this.direccion = ""; 
    this.estado = "";
    this.formato = "";
    this.municipio = "";
    this.razon_social = "";
    this.sucursal = "";
    this.nombre_comercial = "";
    this.dirección_fiscal = "";
    this.RFC = "";
    this.telefono = "";
    this.nombre_contacto = "";
    this.mail_envio_constancias = "";
    this.mail_cco = "";
    this.notas = "";
    
}

}
class Participante { 
  String determinante, sucursal, nombre, apellidos,curp,area_puesto,area_tematica; 
  boolean aprovado;   
  
  public Participante(){
      this.determinante = "";
      this.apellidos="";
      this.aprovado=false;
      this.area_puesto="";
      this.area_tematica="";
      this.curp="";
      this.nombre="";
      this.sucursal="";
  }
} 

class Curso{
    String nombre_empresa, nombre_curso, nombre_instructor,horas_texto,fecha_texto_diploma;
    String razon_social,rfc_empresa, fecha_certificado, capacitador, uCapacitadora, registro_jorge,registro_coco,registro_manuel;
    int horas;
    boolean walmart;
    Date fecha_inicio,fecha_termino;
    
    public Curso() {
    
        this.horas=0;
        this.horas_texto="";
        this.nombre_curso="";
        this.nombre_empresa="";
        this.nombre_instructor="";
        this.fecha_inicio = new Date();
        this.fecha_termino= new Date();
        this.fecha_texto_diploma = "";
        this.fecha_certificado = "";
        this.razon_social = "";
        this.rfc_empresa = "";
        this.uCapacitadora = "";
        this.capacitador = "";
        this.registro_coco = "DPC-ENL-I-103_2015";
        this.registro_manuel = "DPC-ENL-I-056_2016";
        this.registro_jorge = "DPC-ENL-CE-002/2016";
        this.walmart = true;
    }
}

class netoCustomException extends Exception
{
  public netoCustomException(String message)
  {
    super(message);
  }
}

public class Cdiscisa {

    /**
     * @param args the command line arguments
     * @throws java.io.IOException
     * @throws org.apache.poi.openxml4j.exceptions.InvalidFormatException
     */
    
    public static void main(String[] args) throws IOException, InvalidFormatException, Exception {
        
        
        
        //directorioPath, destFolder, listaPath,unidadCapacitadora,instructor,chkDip,chkDipFirma,chkDipLogo,chkConst,chkConstFirma,chkConstLogo,chkDC3a,chkDC3Firmaa,chkDC3Logoa,chkCompilado
        //Workbook wbLista = WorkbookFactory.create(new File("src/files/FormatoCertificado.xlsx"));
        //Workbook wbDirectorio = WorkbookFactory.create(new File("src/files/DirectorioBAE.xlsx"));
        Workbook wbDirectorio = null, wbLista;
        
        //file = (BufferedInputStream)cdiscisa.Cdiscisa.class.getClassLoader().getResourceAsStream("files/DC3_base_firma.pdf");
       
        //JOptionPane.showMessageDialog(null, "directorio " + args[0] +  " Lista : " + args[2]);
        
        if (args[0].isEmpty()){
            try{ 
                InputStream is = new FileInputStream("files/DirectorioBAE.xlsx");
                wbDirectorio = WorkbookFactory.create(is);
            } catch (Exception ex){
                JOptionPane.showMessageDialog(null, "Error al cargar el Directorio de Sucursales \ndirectorio: " + String.valueOf(wbDirectorio) + "\n" + ex.toString());
                return;
            }
        } else {
            try{
                wbDirectorio = WorkbookFactory.create(new File(args[0])); 
            } catch (Exception ex){
                JOptionPane.showMessageDialog(null, "Hubo un error al cargar el Directorio, intenta de nuevo \nDirectorio: " + args[2] + "\n" + ex.toString());
                return;
            }            
        }
        
        wbLista = null;
        
        if (args[2].isEmpty()){
            JOptionPane.showMessageDialog(null, "Es necesario proporcionar la lista de participates de este curso");
            return;            
        } else {
            try{
                wbLista = WorkbookFactory.create(new File(args[2])); 
            } catch (Exception ex){
                JOptionPane.showMessageDialog(null, "Hubo un error al cargar la Lista de participantes, intenta de nuevo \nLista: " + args[2] + "\n" + ex.toString());
                return;
            }                       
        } 

        String nameRegistro;
        
        if (args[15].isEmpty()){
            nameRegistro = "files/RegistroPCJorge Razon2016.pdf";
        } else {        
            nameRegistro = args[15];
        }

        if (args[16].isEmpty()){
            JOptionPane.showMessageDialog(null, "Es necesario proporcionar la lista autógrafa de este curso.");
            return;
        }
        
        ArrayList <Directorio> listaDirectorio = null;
        Map<String, String> abreviaturas = new HashMap<>();

        ArrayList <Participante> listaParticipantes = null;
        Curso c = null;
      
         
        Workbook wbAbrev = null;
        try {
            
            //ClassLoader classLoader = Thread.currentThread().getContextClassLoader();
            //File file = new File(classLoader.getResource("files/abrevcursos.xlsx").toString());
           // BufferedInputStream file= (BufferedInputStream) cdiscisa.Cdiscisa.class.getClassLoader().getClass().getResourceAsStream("files/abrev_cursos.xlsx");
           // File file = new File(cdiscisa.Cdiscisa.class.getClassLoader().getResource("files/abrevcursos.xlsx").getFile());
            //wbAbrev = WorkbookFactory.create(new File(cdiscisa.Cdiscisa.class.getClassLoader().getResource("files/abrev_cursos.xlsx").toString()));
            //InputStream is = new FileInputStream("files/abrevcursos.xlsx");
            InputStream is2 = cdiscisa.Cdiscisa.class.getClassLoader().getResourceAsStream("files/abrevcursos.xlsx");
   
            wbAbrev = WorkbookFactory.create(is2);
                   
        } catch (EncryptedDocumentException EDex){
            JOptionPane.showMessageDialog(null,EDex.getMessage() + "\n" + EDex);
        } catch (IOException ioex){
            JOptionPane.showMessageDialog(null,ioex.getMessage());
        } catch (InvalidFormatException IFEx){
            JOptionPane.showMessageDialog(null,IFEx.getMessage());
        }catch(Exception ex) {
            JOptionPane.showMessageDialog(null,ex.getMessage());
        
        }
         //Workbook wbAbrev = WorkbookFactory.create(file);
        
        try{            
            abreviaturas = llenarAbreviaturas(wbAbrev);
            
        } catch(Exception ex){
            JOptionPane.showMessageDialog(null, "Hubo un error leyendo el archivo de abreviaturas, es posible que contenga celdas faltantes o mal formadas");
            return;
        }
        
        try{
            listaDirectorio = llenarDirectorio(wbDirectorio);
        } catch(Exception ex){
            JOptionPane.showMessageDialog(null, "Hubo un error leyendo el archivo de directorio, es posible que contenga celdas faltantes o mal formadas");
            return;
        }
        try{
            c = llenarCurso(wbLista,args[3],args[4]);
        } catch(Exception ex){
            JOptionPane.showMessageDialog(null, "Hubo un error leyendo la infomacion de los cursos, es posible que contenga celdas faltantes o mal formadas");
            return;
        }
        try{
            listaParticipantes = llenarParticipantes(wbLista);
        } catch(Exception ex){
            JOptionPane.showMessageDialog(null, "Hubo un error leyendo el archivo la Lista de Participantes, es posible que contenga celdas faltantes o mal formadas");
            return;
        }
        
        Map<String,String> dosc = new HashMap<>();
        
        if (args[5].equalsIgnoreCase("true")){ //check if diploma checkbox is checked    
            imprimirDiplomas_main(listaParticipantes, c, listaDirectorio, args[6], args[7], args[1], dosc, args[4], abreviaturas);    // args 6 y 7 son firma y logo respectivamente      
        }
        
        if (args[8].equalsIgnoreCase("true")){            
            imprimirConstancias(listaParticipantes, c, listaDirectorio, args[9], args[10], args[1], dosc, args[4], abreviaturas);
        }
        
        if (args[11].equalsIgnoreCase("true")){            
            imprimirDC3(listaParticipantes,c, args[12], args[13], args[1], dosc, abreviaturas);
        }
        
        if (args[14].equalsIgnoreCase("true")){            
            mergeFiles(dosc,listaParticipantes,listaDirectorio, nameRegistro, args[16]);
        }
        
        JOptionPane.showMessageDialog(null, "Los documentos se han generado exitosamente");

    }
    private static Map<String, String> llenarAbreviaturas(Workbook wbAbrev){
        Map<String, String> abreviaturas = new HashMap<>();        
        Sheet wbListaSheet = wbAbrev.getSheetAt(0);
        Iterator<Row> rowIterator = wbListaSheet.iterator();
        Row row = null; 
        
        while (rowIterator.hasNext()){
            row = rowIterator.next();
            if(row.getCell(0) == null || row.getCell(0).toString().isEmpty())
            {break;} else{
                abreviaturas.put(row.getCell(0).getStringCellValue().trim(), row.getCell(1).getStringCellValue().trim());
            }            
        }
        
        return abreviaturas;
        
    }
    private static Curso llenarCurso (Workbook wbLista, String unidadCapacitadora, String instructor) throws Exception{
        
        Sheet wbListaSheet = wbLista.getSheetAt(0);
        
        Curso c = new Curso();
        
        if (!(wbListaSheet.getRow(2) == null || wbListaSheet.getRow(2).getCell(4) == null || wbListaSheet.getRow(2).getCell(4).getStringCellValue().isEmpty()) ){
            try {
                c.nombre_empresa = wbListaSheet.getRow(2).getCell(4).getStringCellValue();
            } catch (Exception ex){
                JOptionPane.showMessageDialog(null, "El nombre de la empresa en la Lista de Participantes parece tener datos no válidos"); 
                throw new netoCustomException("Error al leer los datos del curso");
            }
            
        }
        
        if (!(wbListaSheet.getRow(4) == null || wbListaSheet.getRow(4).getCell(4) == null || wbListaSheet.getRow(4).getCell(4).getStringCellValue().isEmpty()) ) {
            try{
                c.nombre_curso = wbListaSheet.getRow(4).getCell(4).getStringCellValue();
            } catch (Exception ex){
                JOptionPane.showMessageDialog(null, "El nombre de el curso Lista de Participantes parece tener datos no válidos"); 
                throw new netoCustomException("Error al leer los datos del curso");
            }
        }
        
        if (!(wbListaSheet.getRow(6) == null || wbListaSheet.getRow(6).getCell(4) == null || wbListaSheet.getRow(6).getCell(4).getStringCellValue().isEmpty()) ){
            try{
                c.nombre_instructor = wbListaSheet.getRow(6).getCell(4).getStringCellValue();
            }catch (Exception ex){
                JOptionPane.showMessageDialog(null, "El nombre de el instructor Lista de Participantes parece tener datos no válidos"); 
                throw new netoCustomException("Error al leer los datos del curso");
            }
        }
        
        if (!(wbListaSheet.getRow(8) == null || wbListaSheet.getRow(8).getCell(4) == null || wbListaSheet.getRow(8).getCell(4).getStringCellValue().isEmpty())  ){
            try{
                c.horas_texto = wbListaSheet.getRow(8).getCell(4).getStringCellValue();
            }catch (Exception ex){
                JOptionPane.showMessageDialog(null, "La casilla de Horas en la Lista de Participantes parece tener datos no válidos"); 
                throw new netoCustomException("Error al leer los datos del curso");
            }
        }
        
        if (!(wbListaSheet.getRow(11) == null || wbListaSheet.getRow(11).getCell(4) == null || wbListaSheet.getRow(11).getCell(4).getStringCellValue().isEmpty())  ){
            try{
                c.razon_social = wbListaSheet.getRow(11).getCell(4).getStringCellValue();
                if (!c.razon_social.equalsIgnoreCase("NUEVA WAL‐MART DE MEXICO S DE RL DE C.V.") && !c.razon_social.equalsIgnoreCase("NUEVA WAL-MART DE MEXICO S DE RL DE C.V.")){
                    c.walmart = false;
                }
            }catch (Exception ex){
                JOptionPane.showMessageDialog(null, "La casilla de Razón Social de la empresa en la Lista de Participantes parece tener datos no válidos"); 
                throw new netoCustomException("Error al leer los datos del curso");
            }         
        }
        
        if (!(wbListaSheet.getRow(13) == null || wbListaSheet.getRow(13).getCell(4) == null || wbListaSheet.getRow(13).getCell(4).getStringCellValue().isEmpty())  ){
            try{
                c.rfc_empresa = wbListaSheet.getRow(13).getCell(4).getStringCellValue();
            }catch (Exception ex){
                JOptionPane.showMessageDialog(null, "La casill de RFC de la empresa en Lista de Participantes parece tener datos no válidos"); 
                throw new netoCustomException("Error al leer los datos del curso");
            }         
        } else {        
            JOptionPane.showMessageDialog(null, "El RFC de la empresa no puede estar vacio"); 
            throw new netoCustomException("Error al leer los datos del curso");
        }
        
        if (!(wbListaSheet.getRow(15) == null || wbListaSheet.getRow(15).getCell(4) == null || wbListaSheet.getRow(15).getCell(4).getStringCellValue().isEmpty())  ){
           try{
               c.fecha_certificado = wbListaSheet.getRow(15).getCell(4).getStringCellValue();
           }catch (Exception ex){
                JOptionPane.showMessageDialog(null, "La casilla de la fecha de certificado de la empresa en la Lista de Participantes parece tener datos no válidos"); 
                throw new netoCustomException("Error al leer los datos del curso");
            }               
        }
        
        if (!(wbListaSheet.getRow(17) == null || wbListaSheet.getRow(17).getCell(4) == null || wbListaSheet.getRow(17).getCell(4).getStringCellValue().isEmpty())  ){
            try{
                c.fecha_texto_diploma = wbListaSheet.getRow(17).getCell(4).getStringCellValue();
            }catch (Exception ex){
                JOptionPane.showMessageDialog(null, "La casilla de la fecha para diploma de la empresa en la Lista de Participantes parece tener datos no válidos"); 
                throw new netoCustomException("Error al leer los datos del curso");
            }               
        }
        
        if (!unidadCapacitadora.isEmpty()) {
            c.uCapacitadora = unidadCapacitadora;
        }
        
        if (!instructor.isEmpty()){
            c.capacitador = instructor;
        }
        
        Calendar cal = Calendar.getInstance();
        
        if (!(wbListaSheet.getRow(4) == null || wbListaSheet.getRow(4).getCell(6) == null)){
            if (wbListaSheet.getRow(4).getCell(6).getCellType() == 1){
                cal.set(Calendar.DAY_OF_MONTH, Integer.parseInt(wbListaSheet.getRow(4).getCell(6).getStringCellValue()));                
            } else{
                cal.set(Calendar.DAY_OF_MONTH, (int)wbListaSheet.getRow(4).getCell(6).getNumericCellValue());
            }
         }
        if (!(wbListaSheet.getRow(6) == null || wbListaSheet.getRow(6).getCell(6) == null)){
            cal.set(Calendar.MONTH, Integer.parseInt(wbListaSheet.getRow(6).getCell(6).getStringCellValue())-1);
        }
        if (!(wbListaSheet.getRow(8) == null || wbListaSheet.getRow(8).getCell(6) == null)){
            cal.set(Calendar.YEAR, (int)wbListaSheet.getRow(8).getCell(6).getNumericCellValue());
        }
        
        c.fecha_inicio  = cal.getTime();
        
        if (!(wbListaSheet.getRow(4) == null || wbListaSheet.getRow(4).getCell(7) == null)){
            if (wbListaSheet.getRow(4).getCell(7).getCellType() == 1){
                cal.set(Calendar.DAY_OF_MONTH, Integer.parseInt(wbListaSheet.getRow(4).getCell(7).getStringCellValue()));
            }else{
                cal.set(Calendar.DAY_OF_MONTH, (int)wbListaSheet.getRow(4).getCell(7).getNumericCellValue());
            }
         }
        if (!(wbListaSheet.getRow(6) == null || wbListaSheet.getRow(6).getCell(7) == null)){
            cal.set(Calendar.MONTH, Integer.parseInt(wbListaSheet.getRow(6).getCell(7).getStringCellValue())-1);
        }
        if (!(wbListaSheet.getRow(8) == null || wbListaSheet.getRow(8).getCell(7) == null)){
            cal.set(Calendar.YEAR, (int)wbListaSheet.getRow(8).getCell(7).getNumericCellValue());
        }
        
        c.fecha_termino = cal.getTime();
        
        return c;
    }
    private static ArrayList <Participante> llenarParticipantes (Workbook wbLista) throws Exception{
        
        ArrayList <Participante> listaParticipantes = new ArrayList <>();
        Sheet wbListaSheet = wbLista.getSheetAt(0);
        Iterator<Row> rowIterator = wbListaSheet.iterator();
        
        while(rowIterator.hasNext()){
        
            Row row = rowIterator.next();
            try{
            if (row.getCell(2) != null && row.getCell(2).getStringCellValue().equalsIgnoreCase("# Det.") ){
                break;
            }
            }catch(Exception ex){
                JOptionPane.showMessageDialog(null, "Error leyendo la columna Determinante del archivo Excel de Lista de participantes ");
            }
        }
        
        while(rowIterator.hasNext()){
        
            Row row = rowIterator.next();
            
            row.getCell(2).setCellType(Cell.CELL_TYPE_STRING);
            
            if (row.getCell(2) == null || row.getCell(2).getStringCellValue().isEmpty() )
            {break;}
            
            Participante p = new Participante();
            
            if (row.getCell(2) != null && row.getCell(2).getCellType() != Cell.CELL_TYPE_BLANK && row.getCell(2).getCellType() != Cell.CELL_TYPE_ERROR){                    
                try{
                    p.determinante = row.getCell(2).getStringCellValue().trim();
                }catch(Exception ex){
                    JOptionPane.showMessageDialog(null, "Error leyendo determinante del archivo Excel de Lista de participantes  ");
                }
            }
            
            if (row.getCell(3) != null && row.getCell(3).getCellType() != Cell.CELL_TYPE_BLANK && row.getCell(3).getCellType() != Cell.CELL_TYPE_ERROR){   
                try{
                p.sucursal = row.getCell(3).getStringCellValue().trim();
                }catch(Exception ex){
                    JOptionPane.showMessageDialog(null, "Error leyendo la sucursal del archivo Excel de Lista de participantes  ");
                }
            }
            
            if (row.getCell(4) != null && row.getCell(4).getCellType() != Cell.CELL_TYPE_BLANK && row.getCell(4).getCellType() != Cell.CELL_TYPE_ERROR){   
                try{
                p.nombre = row.getCell(4).getStringCellValue().trim();
                }catch(Exception ex){
                    JOptionPane.showMessageDialog(null, "Error leyendo la columna Nombre del archivo Excel de Lista de participantes  ");
                }
            }
            
            if (row.getCell(5) != null && row.getCell(5).getCellType() != Cell.CELL_TYPE_BLANK && row.getCell(5).getCellType() != Cell.CELL_TYPE_ERROR){ 
                try{
                p.apellidos = row.getCell(5).getStringCellValue().trim();
                }catch(Exception ex){
                    JOptionPane.showMessageDialog(null, "Error leyendo la columna Apellidos del archivo Excel de Lista de participantes  ");
                }
            }
            
            if (row.getCell(6) != null && row.getCell(6).getCellType() != Cell.CELL_TYPE_BLANK && row.getCell(6).getCellType() != Cell.CELL_TYPE_ERROR){ 
                try{
                p.curp = row.getCell(6).getStringCellValue().trim();
                }catch(Exception ex){
                    JOptionPane.showMessageDialog(null, "Error leyendo la columna CURP del archivo Excel de Lista de participantes  ");
                }
            }
            
            if (row.getCell(7) != null && row.getCell(7).getCellType() != Cell.CELL_TYPE_BLANK && row.getCell(7).getCellType() != Cell.CELL_TYPE_ERROR){ 
                try{
                p.area_puesto = row.getCell(7).getStringCellValue().trim();
                }catch(Exception ex){
                    JOptionPane.showMessageDialog(null, "Error leyendo la columna Area Puesto del archivo Excel de Lista de participantes  ");
                }
            }
            
            if (row.getCell(8) != null && row.getCell(8).getCellType() != Cell.CELL_TYPE_BLANK && row.getCell(8).getCellType() != Cell.CELL_TYPE_ERROR){ 
                try{
                p.area_tematica = row.getCell(8).getStringCellValue().trim();
                }catch(Exception ex){
                    JOptionPane.showMessageDialog(null, "Error leyendo la columna Area Tematica del archivo Excel de Lista de participantes  ");
                }
            }
            
            p.aprovado = false;
            if (row.getCell(9) != null && row.getCell(9).getCellType() != Cell.CELL_TYPE_BLANK && row.getCell(9).getCellType() != Cell.CELL_TYPE_ERROR && row.getCell(9).getStringCellValue().equalsIgnoreCase("Aprobado")){                    
                try{
                    p.aprovado = true;
                }catch(Exception ex){
                    JOptionPane.showMessageDialog(null, "Error leyendo la columna Aprobado del archivo Excel de Lista de participantes  ");
                }
            }
            
            listaParticipantes.add(p);
            
        }
        
        
        
        return listaParticipantes;
 // method body
}
    private static ArrayList <Directorio> llenarDirectorio (Workbook wbDirectorio) throws netoCustomException, Exception{
        ArrayList <Directorio> listaDirectorio = new ArrayList <>();
        Sheet wbListaSheet = wbDirectorio.getSheetAt(0);
        Iterator<Row> rowIterator = wbListaSheet.iterator();
        
        Row row = null; 
        Directorio d = null;
        boolean walmart = true;
        
        if (rowIterator.hasNext()){
            row = rowIterator.next();
            if (row.getCell(0) != null && !row.getCell(0).toString().equalsIgnoreCase("Det.") )
            {
                walmart = false;
            } 
            
        }
        
        while(rowIterator.hasNext()){
        
            row = rowIterator.next();
            
            try{
                
                if (walmart && (row.getCell(0) == null || row.getCell(0).toString().isEmpty()) ){
                    break;
                } else if(row.getCell(0) == null || row.getCell(0).toString().isEmpty()){
                    continue;                                             
                }
                
                
                d = new Directorio();

                if (row.getCell(0) != null){
                    row.getCell(0).setCellType(Cell.CELL_TYPE_STRING);
                    d.determinante = row.getCell(0).getStringCellValue().trim();
                }

                if (row.getCell(1) != null){   
                    if (walmart){
                        d.formato= row.getCell(1).getStringCellValue().trim();
                    } else
                    {
                        d.razon_social= row.getCell(1).getStringCellValue().trim();
                    }
                }

                if (row.getCell(2) != null){    
                    if (walmart){
                        d.unidad = row.getCell(2).getStringCellValue().trim();
                    }else{
                        d.sucursal = row.getCell(2).getStringCellValue().trim();
                    }
                }

                if (row.getCell(3) != null){                     
                    d.estado = row.getCell(3).getStringCellValue().trim();
                }
                if (row.getCell(4) != null){                     
                    d.municipio = row.getCell(4).getStringCellValue().trim();
                }
                if (row.getCell(5) != null){                     
                    d.direccion = row.getCell(5).getStringCellValue().trim();
                }
                
                if (!walmart){
                    if (row.getCell(6) != null){                     
                        d.nombre_comercial = row.getCell(6).getStringCellValue().trim();
                    }
                    if (row.getCell(7) != null){                     
                        d.dirección_fiscal = row.getCell(7).getStringCellValue().trim();
                    }
                    if (row.getCell(8) != null){                     
                        d.RFC = row.getCell(8).getStringCellValue().trim();
                    }
                    if (row.getCell(9) != null){                     
                        d.telefono = row.getCell(9).getStringCellValue().trim();
                    }
                    if (row.getCell(10) != null){                     
                        d.nombre_contacto = row.getCell(10).getStringCellValue().trim();
                    }
                    if (row.getCell(11) != null){                     
                        d.mail_envio_constancias = row.getCell(11).getStringCellValue().trim();
                    }
                    if (row.getCell(12) != null){                     
                        d.mail_cco = row.getCell(12).getStringCellValue().trim();
                    }
                    if (row.getCell(13) != null){                     
                        d.notas = row.getCell(13).getStringCellValue().trim();
                    }
                }
            
            }catch(Exception ex){
                JOptionPane.showMessageDialog(null, "Error leyendo la determinante " + d.determinante + " " + d.unidad); 
            }
            listaDirectorio.add(d);
            
        }
        
        return listaDirectorio;
    }
    private static void imprimirDiplomas(ArrayList <Participante> listaParticipantes, Curso c, Directorio d, String chkDipFirma, String chkDipLogo, String savePath, Map<String,String> dosc, String instructor, Map<String, String> abreviaturas) throws IOException{
        
          
        ListIterator <Participante> it = listaParticipantes.listIterator();
        Participante p1 = null;
        Participante p2 = null;
        String abrev_curso = "";


        // Create a document and add a page to it
        PDDocument document = new PDDocument();
        PDDocument documentSingle;
        
        InputStream file = null;
        
        BufferedImage logo = null;
        BufferedImage firma = null;
        
        try{    
            
            //logo = new File(cdiscisa.Cdiscisa.class.getClassLoader().getResource("files/logo.png").getFile());
            //logo = cdiscisa.Cdiscisa.class.getClassLoader().getResourceAsStream("files/logo.png");
            logo = ImageIO.read(cdiscisa.Cdiscisa.class.getClassLoader().getResourceAsStream("files/logo.png"));
            //logo = cdiscisa.Cdiscisa.class.getClassLoader().getResourceAsStream("files/logo.png");
            
             if(instructor.equalsIgnoreCase("Ing. Jorge Antonio Razón Gutierrez")){
                 //firma = new File(cdiscisa.Cdiscisa.class.getClassLoader().getResource("files/firmaCoco.png").getFile());
                 firma = ImageIO.read(cdiscisa.Cdiscisa.class.getClassLoader().getResourceAsStream("files/firmaCoco.png"));
             } else if (instructor.equalsIgnoreCase("Manuel Anguiano Razón")){
                 firma = ImageIO.read(cdiscisa.Cdiscisa.class.getClassLoader().getResourceAsStream("files/firmaManuel.png"));
                 //firma = new File(cdiscisa.Cdiscisa.class.getClassLoader().getResource("files/firmaManuel.png").getFile());
             } else {
                 firma = ImageIO.read(cdiscisa.Cdiscisa.class.getClassLoader().getResourceAsStream("files/firmaJorge.png"));
                 //firma = new File(cdiscisa.Cdiscisa.class.getClassLoader().getResource("files/firmaJorge.png").getFile());
            }
             
        }catch (Exception ex){
            JOptionPane.showMessageDialog(null, "Error al cargar la imagen del logo o la firma \nfile: " + String.valueOf(logo) + "\n" + String.valueOf(firma) + "\n" + ex.toString());
        }
        
        PDImageXObject firmaObject = null;
        PDImageXObject logoObject = null;
        
        try{
            if (chkDipLogo.equalsIgnoreCase("true")){
                //logoObject = PDImageXObject.createFromFile(logo, document);                
                logoObject = LosslessFactory.createFromImage(document, logo);               
            }
            
            if (chkDipFirma.equalsIgnoreCase("true")){
               // firmaObject = PDImageXObject.createFromFile(firma, document);
               firmaObject = LosslessFactory.createFromImage(document, firma);
            }
            
        }catch(Exception ex){
            JOptionPane.showMessageDialog(null, "Error al crear objetos de logo o firma \nfile: " + String.valueOf(logoObject) + "\n" + String.valueOf(firmaObject) + "\n" + ex.toString());
        }
        

        try{
            file = cdiscisa.Cdiscisa.class.getClassLoader().getResourceAsStream("files/n_diploma_simple_vacio_nf_nl_nr.pdf");
        } catch (Exception ex){
            JOptionPane.showMessageDialog(null, "Error al cargar el el diploma single base. \ndiploma: files/n_diploma_simple_vacio_nf_nl_nr.pdf \nfile: " + String.valueOf(file) + "\n" + ex.toString());
        }
                
        documentSingle = PDDocument.load(file);
        
        PDPage pageSingle = (PDPage)documentSingle.getDocumentCatalog().getPages().get(0);  
        COSDictionary pageDictSingle = pageSingle.getCOSObject();
        COSDictionary newPageSingleDict = new COSDictionary(pageDictSingle);
        PDPage templatePageSingle = new PDPage(newPageSingleDict);
        
        // Create a document and add a page to it
        
        PDDocument documentDoble;
        
        InputStream file2 = null;

        try{
            file2 = cdiscisa.Cdiscisa.class.getClassLoader().getResourceAsStream("files/n_diploma_doble_vacio_nf_nl_nr.pdf");
        } catch (Exception ex){
            JOptionPane.showMessageDialog(null, "Error al cargar el el diploma doble base. \ndiploma: n_diploma_doble_vacio_nf_nl_nr.pdf\nfile: " + String.valueOf(file2) + "\n" + ex.toString());
        }
        
        documentDoble = PDDocument.load(file2);
        
        PDPage pageDoble = (PDPage)documentDoble.getDocumentCatalog().getPages().get(0);  
        COSDictionary pageDobleDict = pageDoble.getCOSObject();
        COSDictionary newPageDobleDict = new COSDictionary(pageDobleDict);
       
        PDPage templatePageDoble = new PDPage(newPageDobleDict);
        
        InputStream isFont1 = null,isFont2 = null,isFont3 = null;
        
        try{
        isFont1 = cdiscisa.Cdiscisa.class.getClassLoader().getResourceAsStream("files/Calibri.ttf");    
        isFont2 = cdiscisa.Cdiscisa.class.getClassLoader().getResourceAsStream("files/CalibriBold.ttf");        
        isFont3 = cdiscisa.Cdiscisa.class.getClassLoader().getResourceAsStream("files/Pristina.ttf");      
        } catch (Exception ex){
            JOptionPane.showMessageDialog(null, "Error al cargar el una fuente \nisFont1: " + String.valueOf(isFont1) + "\nisFont2: " + String.valueOf(isFont2) +  "\nisFont3: " + String.valueOf(isFont3) + "\n" + ex.toString());
        }
        PDFont calibri =  null;
        PDFont calibriBold = null;
        PDFont pristina =  null;

        calibri = PDType0Font.load(document, isFont1);
        calibriBold = PDType0Font.load(document,isFont2);
        pristina = PDType0Font.load(document, isFont3);
        
        if(listaParticipantes.size() % 2 == 0 && listaParticipantes.size() >= 2){
            
            while (it.hasNext()){
                p1 = it.next();
                p2 = it.next();
                
                imprimirDiplomaDoble(p1,p2,c,d,document,templatePageDoble, calibri,calibriBold,pristina, logoObject, firmaObject, instructor);
            } 
        } else {
            if (listaParticipantes.size() > 1){
                //Lista es impar y contiene mas de 2 participantes.
                while (it.hasNext()){
                    
                    if (it.nextIndex() == listaParticipantes.size() - 1){
                        p1 = it.next();
                        imprimirDiplomaArriba(p1,c,d,document,templatePageSingle,calibri,calibriBold,pristina, logoObject, firmaObject, instructor);
                        break;
                    }
                    
                    p1 = it.next();
                    p2 = it.next();
                
                    imprimirDiplomaDoble(p1,p2,c,d,document,templatePageDoble,calibri,calibriBold,pristina, logoObject, firmaObject, instructor);
                }          
            } else if (listaParticipantes.size() == 1){
                p1 = it.next();
                imprimirDiplomaArriba(p1,c,d,document,templatePageSingle,calibri,calibriBold,pristina,logoObject, firmaObject, instructor);
            }  
            
        }
        
        Format formatter = new SimpleDateFormat("ddMMMYYYY", new Locale("es","MX"));
        String formatedDate = formatter.format(c.fecha_inicio);
        
        String abrev = abreviaturas.get(c.nombre_curso);
        
        if (c.walmart){
            document.save(savePath + File.separator + "Diplomas_" + d.formato + "_" + d.unidad + "_" + d.determinante + "_" + abrev + "_" + formatedDate + ".pdf");
            document.close();        
            dosc.put(savePath + File.separator + "Diplomas_" + d.formato + "_" + d.unidad + "_" + d.determinante + "_" + abrev + "_" + formatedDate + ".pdf",d.determinante);
        }else{
            document.save(savePath + File.separator + "Diplomas_" + d.determinante + "_" + abrev + "_" + formatedDate + ".pdf");
            document.close();        
            dosc.put(savePath + File.separator + "Diplomas_" + d.determinante + "_" + abrev + "_" + formatedDate + ".pdf",d.determinante);
        }
    }
    private static void imprimirDiplomaArriba(Participante p1, Curso c, Directorio d, PDDocument document, PDPage page, PDFont calibri, PDFont calibriBold, PDFont pristina, PDImageXObject logoObject, PDImageXObject firmaObject, String instructor) throws IOException{
        
       
        
        COSDictionary pageDict = page.getCOSObject();
        COSDictionary newPageDict = new COSDictionary(pageDict);

        PDPage newPage = new PDPage(newPageDict);
        document.addPage(newPage);

        // Start a new content stream which will "hold" the to be created content
        PDPageContentStream contentStream = new PDPageContentStream(document, newPage, true, true);

        float pageWidth = newPage.getMediaBox().getWidth();
        //float pageHeight = newPage.getMediaBox().getHeight();
        
        //System.out.println("pageWidth: " + pageWidth + "\npageHeight: " + pageHeight);

        // Print Name
        contentStream.beginText();
        contentStream.setFont( pristina, 28 );
        //contentStream.setNonStrokingColor(0,112,192);
        contentStream.setNonStrokingColor(0,128,0);
        
        float nameWidth = pristina.getStringWidth(p1.nombre + " " + p1.apellidos) /1000 * 28;
        
        //System.out.println(p1.nombre + " " + p1.apellidos + "::" + nameWidth);
        
        float xPosition;
        if (nameWidth < 470){
            xPosition = (pageWidth - nameWidth)/2;
            contentStream.setTextMatrix(new Matrix(1,0,0,1,xPosition,641 ));           
            contentStream.showText(p1.nombre + " " + p1.apellidos);
        } else{
            contentStream.setFont( pristina, 22 );
            nameWidth = pristina.getStringWidth(p1.nombre + " " + p1.apellidos) /1000 * 22;
            xPosition = (pageWidth - nameWidth)/2;
            contentStream.setTextMatrix(new Matrix(1,0,0,1,xPosition,641 ));           
            contentStream.showText(p1.nombre + " " + p1.apellidos);
        }
        
        contentStream.setFont( calibri, 15 );
        contentStream.setNonStrokingColor(Color.BLACK);
        
        nameWidth = calibri.getStringWidth(c.razon_social) /1000 * 15;
        //System.out.println("nameWidth: " + nameWidth);
        xPosition = (pageWidth - nameWidth)/2;
         
        contentStream.setTextMatrix(new Matrix(1,0,0,1,xPosition,615 ));           
        contentStream.showText(c.razon_social);
        
        contentStream.setFont( calibri, 11 );
        contentStream.setNonStrokingColor(Color.BLACK);
         
        contentStream.setTextMatrix(new Matrix(1,0,0,1,255, 593));           
        if (c.walmart){
            contentStream.showText(p1.determinante + " " + d.unidad);
        } else if (d.sucursal.isEmpty() || d.sucursal.equalsIgnoreCase("")){
            contentStream.endText();
            contentStream.addRect(100, 590, 400, 20);
            contentStream.setNonStrokingColor(Color.WHITE);
            contentStream.fill();
            contentStream.beginText();
            contentStream.setNonStrokingColor(Color.BLACK);
        } else{
            contentStream.showText(d.sucursal);
        }
         
        contentStream.setTextMatrix(new Matrix(1,0,0,1,210, (float)538.5));           
        contentStream.showText(c.horas_texto);
        
        contentStream.setTextMatrix(new Matrix(1,0,0,1,394,(float)538.5));           
        contentStream.showText(c.fecha_texto_diploma);
        
        contentStream.setFont( calibriBold, 10 ); 
        
        float nameWidthStroked = calibri.getStringWidth(c.nombre_curso) /1000 * 10;
        //System.out.println("nameWidth: " + nameWidthStroked);
        float strokePosition = (pageWidth - nameWidthStroked)/2;
        contentStream.setTextMatrix(new Matrix(1,0,0,1,strokePosition, 554 ));           
        contentStream.showText(c.nombre_curso);
        
        contentStream.setFont( calibri, 8 ); 
        contentStream.setNonStrokingColor(Color.GRAY);
        
        if (instructor.equalsIgnoreCase("Manuel Anguiano Razón")){
            contentStream.setTextMatrix(new Matrix(1,0,0,1,170, 457 ));
            contentStream.showText("Registro STPS: GIS100219KK8003");
            contentStream.setTextMatrix(new Matrix(1,0,0,1,170, 447 ));
            contentStream.showText("Registro PC: " + c.registro_manuel);
            contentStream.setTextMatrix(new Matrix(1,0,0,1,170, 437 ));
            contentStream.showText("Registro PC: " + c.registro_jorge);
            
        } else if (instructor.equalsIgnoreCase("Ing. Jorge Antonio Razón Gutierrez")){
            contentStream.setTextMatrix(new Matrix(1,0,0,1,170, 457 ));
            contentStream.showText("Registro STPS: GIS100219KK8003");
            contentStream.setTextMatrix(new Matrix(1,0,0,1,170, 447 ));
            contentStream.showText("Registro PC: " + c.registro_coco);
            contentStream.setTextMatrix(new Matrix(1,0,0,1,170, 437 ));
            contentStream.showText("Registro PC: " + c.registro_jorge);
        } else {
            contentStream.setTextMatrix(new Matrix(1,0,0,1,150, 457 ));
            contentStream.showText("Registro STPS: GIS100219KK8003");
            contentStream.setTextMatrix(new Matrix(1,0,0,1,150, 447 ));
            contentStream.showText("Registro STPS: RAGJ610813BIA005");
            contentStream.setTextMatrix(new Matrix(1,0,0,1,150, 437 ));
            contentStream.showText("Registro PC: " + c.registro_jorge);
            contentStream.setTextMatrix(new Matrix(1,0,0,1,270, 447 ));
            contentStream.showText("Registro PC: SPC-COAH-056-2015");
            contentStream.setTextMatrix(new Matrix(1,0,0,1,270, 437));
            contentStream.showText("Registro PC: CGPC-28/6016/026/NL-PS14");
        }
        // "Registro PC: DPC-ENL-CE-002/2015"
        // DPC-ENL-I-103_2015 "Ing. Jorge Antonio Razón Gutierrez"
        // DPC-ENL-I-056_2015 "Manuel Anguiano Razón"
        // "TSI. Jorge Antonio Razón Gil"

        contentStream.endText();
        
        contentStream.setStrokingColor(Color.BLACK);
        contentStream.setLineWidth(1);
        contentStream.moveTo(strokePosition,552);
        contentStream.lineTo(strokePosition + nameWidthStroked + 6, 552);        
        contentStream.stroke();
        
        if (logoObject != null && !logoObject.isEmpty()){
            contentStream.drawImage(logoObject, 451,700,130,65);
        }
        if (firmaObject != null && !firmaObject.isEmpty()){
            contentStream.drawImage(firmaObject, 452,440,110,42);

            contentStream.beginText();
            contentStream.setFont(calibri, 10);
            contentStream.setNonStrokingColor(Color.BLACK);
            nameWidth = calibri.getStringWidth(instructor) /1000 * 10;            
            xPosition = 505 - nameWidth/2;
            
            contentStream.setTextMatrix(new Matrix(1,0,0,1,xPosition, 445 ));           
            contentStream.showText(instructor);
            contentStream.setTextMatrix(new Matrix(1,0,0,1,485, 435 ));           
            contentStream.showText("Instructor");
            contentStream.endText();
        }
        /*
        System.out.println("logoWidth: " + logoObject.getWidth());
        System.out.println("logoHeight: " + logoObject.getHeight());
        System.out.println("firmaWidth: " + firmaObject.getWidth());
        System.out.println("firmaHeight: " + firmaObject.getHeight());
        */
        
        
        
        /*
        contentStream.addRect(50, 750, 500, 100);
        contentStream.setNonStrokingColor(Color.WHITE);
        contentStream.fill();
        contentStream.drawImage(logo, 430,700,150,75);
        contentStream.drawImage(firma, 93,239,72,29);
        */
         
        // Make sure that the content stream is closed:
        contentStream.close();
        
        // Save the results and ensure that the document is properly closed:
        


    }
    private static void imprimirDiplomaDoble(Participante p1, Participante p2, Curso c, Directorio d, PDDocument document, PDPage page, PDFont calibri, PDFont calibriBold, PDFont pristina, PDImageXObject logoObject, PDImageXObject firmaObject, String instructor) throws IOException{
        
       
        
        COSDictionary pageDict = page.getCOSObject();
        COSDictionary newPageDict = new COSDictionary(pageDict);

        PDPage newPage = new PDPage(newPageDict);
        document.addPage(newPage);
        
        // Start a new content stream which will "hold" the to be created content
        PDPageContentStream contentStream = new PDPageContentStream(document, newPage, true, true);
                
        float pageWidth = newPage.getMediaBox().getWidth();
        //float pageHeight = newPage.getMediaBox().getHeight();
        
        //System.out.println("pageWidth: " + pageWidth + "\npageHeight: " + pageHeight);

        // Print UP Side
        
        contentStream.beginText();
        contentStream.setFont( pristina, 28 );
        //contentStream.setNonStrokingColor(0,112,192);
        contentStream.setNonStrokingColor(0,128,0);
        
        float nameWidth = pristina.getStringWidth(p1.nombre + " " + p1.apellidos) /1000 * 28;
        
        //System.out.println(p1.nombre + " " + p1.apellidos + "::" + nameWidth);

        float xPosition;
        if (nameWidth < 470){
            xPosition = (pageWidth - nameWidth)/2;
            contentStream.setTextMatrix(new Matrix(1,0,0,1,xPosition,641 ));           
            contentStream.showText(p1.nombre + " " + p1.apellidos);
        } else{
            contentStream.setFont( pristina, 22 );
            nameWidth = pristina.getStringWidth(p1.nombre + " " + p1.apellidos) /1000 * 22;
            xPosition = (pageWidth - nameWidth)/2;
            contentStream.setTextMatrix(new Matrix(1,0,0,1,xPosition,641 ));           
            contentStream.showText(p1.nombre + " " + p1.apellidos);
        }
        
        contentStream.setFont( calibri, 15 );
        contentStream.setNonStrokingColor(Color.BLACK);
        
        nameWidth = calibri.getStringWidth(c.razon_social) /1000 * 15;
        //System.out.println("nameWidth: " + nameWidth);
        xPosition = (pageWidth - nameWidth)/2;
         
        contentStream.setTextMatrix(new Matrix(1,0,0,1,xPosition,615 ));           
        contentStream.showText(c.razon_social);
        
        contentStream.setFont( calibri, 11 );
        contentStream.setNonStrokingColor(Color.BLACK);
        
        if (c.walmart){
            contentStream.setTextMatrix(new Matrix(1,0,0,1,275, 593)); 
            contentStream.showText(p1.determinante + " " + d.unidad);
        } else if (d.sucursal.isEmpty() || d.sucursal.equalsIgnoreCase("")){
            contentStream.endText();
            contentStream.addRect(100, 590, 400, 20);
            contentStream.setNonStrokingColor(Color.WHITE);
            contentStream.fill();
            contentStream.beginText();
            contentStream.setNonStrokingColor(Color.BLACK);
        } else{
            contentStream.setTextMatrix(new Matrix(1,0,0,1,275, 593));
            contentStream.showText(d.sucursal);
        }
        
        
        contentStream.setTextMatrix(new Matrix(1,0,0,1,210,(float)538.5));           
        contentStream.showText(c.horas_texto);
        
        contentStream.setTextMatrix(new Matrix(1,0,0,1,394,(float)538.5));           
        contentStream.showText(c.fecha_texto_diploma);
        
        contentStream.setFont( calibriBold, 10 ); 
        
        float nameWidthStroked = calibri.getStringWidth(c.nombre_curso) /1000 * 10;
        //System.out.println("nameWidth: " + nameWidthStroked);
        float strokePosition = (pageWidth - nameWidthStroked)/2;
        contentStream.setTextMatrix(new Matrix(1,0,0,1,strokePosition, 554 ));           
        contentStream.showText(c.nombre_curso);
        
        
        contentStream.setFont( calibri, 8 ); 
        contentStream.setNonStrokingColor(Color.GRAY);
        
        if (instructor.equalsIgnoreCase("Manuel Anguiano Razón")){
            contentStream.setTextMatrix(new Matrix(1,0,0,1,170, 457 ));
            contentStream.showText("Registro STPS: GIS100219KK8003");
            contentStream.setTextMatrix(new Matrix(1,0,0,1,170, 447 ));
            contentStream.showText("Registro PC: " + c.registro_manuel);
            contentStream.setTextMatrix(new Matrix(1,0,0,1,170, 437 ));
            contentStream.showText("Registro PC: " + c.registro_jorge);
            
        } else if (instructor.equalsIgnoreCase("Ing. Jorge Antonio Razón Gutierrez")){
            contentStream.setTextMatrix(new Matrix(1,0,0,1,170, 457 ));
            contentStream.showText("Registro STPS: GIS100219KK8003");
            contentStream.setTextMatrix(new Matrix(1,0,0,1,170, 447 ));
            contentStream.showText("Registro PC: " + c.registro_coco);
            contentStream.setTextMatrix(new Matrix(1,0,0,1,170, 437 ));
            contentStream.showText("Registro PC: " + c.registro_jorge);
        } else {
            contentStream.setTextMatrix(new Matrix(1,0,0,1,150, 457 ));
            contentStream.showText("Registro STPS: GIS100219KK8003");
            contentStream.setTextMatrix(new Matrix(1,0,0,1,150, 447 ));
            contentStream.showText("Registro STPS: RAGJ610813BIA005");
            contentStream.setTextMatrix(new Matrix(1,0,0,1,150, 437 ));
            contentStream.showText("Registro PC: " + c.registro_jorge);
            contentStream.setTextMatrix(new Matrix(1,0,0,1,270, 447 ));
            contentStream.showText("Registro PC: SPC-COAH-056-2015");
            contentStream.setTextMatrix(new Matrix(1,0,0,1,270, 437));
            contentStream.showText("Registro PC: CGPC-28/6016/026/NL-PS14");
        }
        
        contentStream.endText();
        
        contentStream.setStrokingColor(Color.BLACK);
        contentStream.setLineWidth(1);  
        contentStream.moveTo(strokePosition,552);
        contentStream.lineTo(strokePosition + nameWidthStroked + 6,552);
        contentStream.stroke();
        
        if (logoObject != null && !logoObject.isEmpty()){
            contentStream.drawImage(logoObject, 451,700,130,65);
        }
        if (firmaObject != null && !firmaObject.isEmpty()){
            contentStream.setStrokingColor(Color.BLACK);
            contentStream.drawImage(firmaObject, 452,440,110,42);
            contentStream.beginText();
            contentStream.setFont(calibri, 10);
            nameWidth = calibri.getStringWidth(instructor) /1000 * 10;            
            xPosition = 505 - nameWidth/2;
            
            contentStream.setTextMatrix(new Matrix(1,0,0,1,xPosition, 445 ));           
            contentStream.showText(instructor);
            contentStream.setTextMatrix(new Matrix(1,0,0,1,485, 435 ));           
            contentStream.showText("Instructor");
            contentStream.endText();
        }
        
        
        
        // Print DOWN Side
        
        contentStream.beginText();
        contentStream.setFont( pristina, 28 );
        //contentStream.setNonStrokingColor(0,112,192);
        contentStream.setNonStrokingColor(0,128,0);
        
        nameWidth = pristina.getStringWidth(p2.nombre + " " + p2.apellidos) /1000 * 28;
        
        //System.out.println(p2.nombre + " " + p2.apellidos + "::" + nameWidth);
       
        if (nameWidth < 470){
            xPosition = (pageWidth - nameWidth)/2;
            contentStream.setTextMatrix(new Matrix(1,0,0,1,xPosition,262 ));           
            contentStream.showText(p2.nombre + " " + p2.apellidos);
        } else{
            contentStream.setFont( pristina, 22 );
            nameWidth = pristina.getStringWidth(p2.nombre + " " + p2.apellidos) /1000 * 22;
            xPosition = (pageWidth - nameWidth)/2;
            contentStream.setTextMatrix(new Matrix(1,0,0,1,xPosition,262 ));           
            contentStream.showText(p2.nombre + " " + p2.apellidos);
        }
        
        
        contentStream.setFont( calibri, 15 );
        contentStream.setNonStrokingColor(Color.BLACK);
        
        nameWidth = calibri.getStringWidth(c.razon_social) /1000 * 15;
        //System.out.println("nameWidth: " + nameWidth);
        xPosition = (pageWidth - nameWidth)/2;
         
        contentStream.setTextMatrix(new Matrix(1,0,0,1,xPosition,235 ));           
        contentStream.showText(c.razon_social);
        
        contentStream.setFont( calibri, 11 );
        contentStream.setNonStrokingColor(Color.BLACK);
               
        if (c.walmart){
            contentStream.setTextMatrix(new Matrix(1,0,0,1,275, (float) 213.4)); 
            contentStream.showText(p2.determinante + " " + d.unidad);
        } else if (d.sucursal.isEmpty() || d.sucursal.equalsIgnoreCase("")){
            contentStream.endText();
            contentStream.addRect(100, 211, 400, 20);
            contentStream.setNonStrokingColor(Color.WHITE);
            contentStream.fill();
            contentStream.beginText();
            contentStream.setNonStrokingColor(Color.BLACK);
        } else{
            contentStream.setTextMatrix(new Matrix(1,0,0,1,275, (float) 213.4)); 
            contentStream.showText(d.sucursal);
        }
        
        
        contentStream.setTextMatrix(new Matrix(1,0,0,1,210,159));           
        contentStream.showText(c.horas_texto);
        
        contentStream.setTextMatrix(new Matrix(1,0,0,1,394,159));           
        contentStream.showText(c.fecha_texto_diploma);
        
        contentStream.setFont( calibriBold, 10 ); 
        
        nameWidthStroked = calibri.getStringWidth(c.nombre_curso) /1000 * 10;
        //System.out.println("nameWidth: " + nameWidthStroked);
        strokePosition = (pageWidth - nameWidthStroked)/2;
        contentStream.setTextMatrix(new Matrix(1,0,0,1,strokePosition, 174 ));           
        contentStream.showText(c.nombre_curso);
        
        
        contentStream.setFont( calibri, 8 ); 
        contentStream.setNonStrokingColor(Color.GRAY);
        
        if (instructor.equalsIgnoreCase("Manuel Anguiano Razón")){
            contentStream.setTextMatrix(new Matrix(1,0,0,1,170, 74 ));
            contentStream.showText("Registro STPS: GIS100219KK8003");
            contentStream.setTextMatrix(new Matrix(1,0,0,1,170, 64 ));
            contentStream.showText("Registro PC: " + c.registro_manuel);
            contentStream.setTextMatrix(new Matrix(1,0,0,1,170, 54 ));
            contentStream.showText("Registro PC: " + c.registro_jorge);
            
        } else if (instructor.equalsIgnoreCase("Ing. Jorge Antonio Razón Gutierrez")){
            contentStream.setTextMatrix(new Matrix(1,0,0,1,170, 74 ));
            contentStream.showText("Registro STPS: GIS100219KK8003");
            contentStream.setTextMatrix(new Matrix(1,0,0,1,170, 64 ));
            contentStream.showText("Registro PC: " + c.registro_coco);
            contentStream.setTextMatrix(new Matrix(1,0,0,1,170, 54 ));
            contentStream.showText("Registro PC: " + c.registro_jorge);
        } else {
            contentStream.setTextMatrix(new Matrix(1,0,0,1,150, 74 ));
            contentStream.showText("Registro STPS: GIS100219KK8003");
            contentStream.setTextMatrix(new Matrix(1,0,0,1,150, 64 ));
            contentStream.showText("Registro STPS: RAGJ610813BIA005");
            contentStream.setTextMatrix(new Matrix(1,0,0,1,150, 54 ));
            contentStream.showText("Registro PC: " + c.registro_jorge);
            contentStream.setTextMatrix(new Matrix(1,0,0,1,270, 64 ));
            contentStream.showText("Registro PC: SPC-COAH-056-2015");
            contentStream.setTextMatrix(new Matrix(1,0,0,1,270, 54));
            contentStream.showText("Registro PC: CGPC-28/6016/026/NL-PS14");
        }
        
        contentStream.endText();
        
        contentStream.setStrokingColor(Color.BLACK);
        contentStream.moveTo(strokePosition,172);
        contentStream.lineTo(strokePosition + nameWidthStroked + 6, 172);
        contentStream.stroke();
        
        if (logoObject != null && !logoObject.isEmpty()){
            contentStream.drawImage(logoObject, 451,320,130,65);
        }
        if (firmaObject != null && !firmaObject.isEmpty()){
            contentStream.drawImage(firmaObject, 452,62,110,42);

            contentStream.beginText();
            contentStream.setFont(calibri, 10);
            contentStream.setStrokingColor(Color.BLACK);
            nameWidth = calibri.getStringWidth(instructor) /1000 * 10;
            
            xPosition = 505 - nameWidth/2;
            
            contentStream.setTextMatrix(new Matrix(1,0,0,1,xPosition, 67 ));
            
            contentStream.showText(instructor);
            contentStream.setTextMatrix(new Matrix(1,0,0,1,485, 57 ));           
            contentStream.showText("Instructor");
            contentStream.endText();
        }
        // Make sure that the content stream is closed:
        contentStream.close();
        
        // Save the results and ensure that the document is properly closed:
        //document.save( "DiplomaSoloTest.pdf");
        //document.close();
    }
    private static void imprimirDC3(ArrayList <Participante> listaParticipantes, Curso c, String chkDC3Firmaa, String chkDC3Logoa, String savePath,  Map<String,String> dosc, Map<String, String> abreviaturas) throws IOException{   
        
        ListIterator <Participante> it = listaParticipantes.listIterator();
        Participante p;
        while(it.hasNext()){
            p = it.next();
            if (p.aprovado){
            imprimirDC3_individual(p,c, chkDC3Firmaa, chkDC3Logoa, savePath, dosc, abreviaturas);
            }
        }  
          
    }
    private static void imprimirDC3_individual(Participante p,Curso c, String chkDC3Firmaa, String chkDC3Logoa, String savePath, Map<String,String> dosc, Map<String, String> abreviaturas) throws IOException {
        
        //JOptionPane.showMessageDialog(null, "entrando a imrpimir individual");
        PDDocument document;
        //BufferedInputStream file;
        InputStream file = null;
        
        try{
            file = cdiscisa.Cdiscisa.class.getClassLoader().getResourceAsStream("files/DC3_blank.pdf");
            //file = cdiscisa.Cdiscisa.class.getClassLoader().getResourceAsStream("files/DC3_base_firma.pdf");
        }catch (Exception ex){
            JOptionPane.showMessageDialog(null, "Error al cargar el una forma DC3 \nfile: " + String.valueOf(file) + "\n" + ex.toString());
        }
        document = PDDocument.load(file);
        
        BufferedImage logo = null;
        BufferedImage firma = null;
        
        try{
             //logo = new File(cdiscisa.Cdiscisa.class.getClassLoader().getResource("files/logo.png").getFile());
             logo = ImageIO.read(cdiscisa.Cdiscisa.class.getClassLoader().getResourceAsStream("files/logo.png"));
             //logo = StreamUtil.stream2file(cdiscisa.Cdiscisa.class.getClassLoader().getResourceAsStream("files/logo.png"));
             if(c.capacitador.equalsIgnoreCase("Ing. Jorge Antonio Razón Gutierrez")){
                 //firma = new File(cdiscisa.Cdiscisa.class.getClassLoader().getResource("files/firmaCoco.png").getFile());
                 firma = ImageIO.read(cdiscisa.Cdiscisa.class.getClassLoader().getResourceAsStream("files/firmaCoco.png"));
             } else if (c.capacitador.equalsIgnoreCase("Manuel Anguiano Razón")){
                 firma = ImageIO.read(cdiscisa.Cdiscisa.class.getClassLoader().getResourceAsStream("files/firmaManuel.png"));
                 //firma = new File(cdiscisa.Cdiscisa.class.getClassLoader().getResource("files/firmaManuel.png").getFile());
             } else {
                 firma = ImageIO.read(cdiscisa.Cdiscisa.class.getClassLoader().getResourceAsStream("files/firmaJorge.png"));
                 //firma = new File(cdiscisa.Cdiscisa.class.getClassLoader().getResource("files/firmaJorge.png").getFile());
            }
             
        }catch (Exception ex){
            JOptionPane.showMessageDialog(null, "Error al cargar la imagen del logo o la firma \nfile: " + String.valueOf(logo) + "\n" + String.valueOf(firma) + "\n" + ex.toString());
        }
        
        PDImageXObject firmaObject = null;
        PDImageXObject logoObject = null;
        
        try{
            if (chkDC3Firmaa.equalsIgnoreCase("true")){
                firmaObject = LosslessFactory.createFromImage(document, firma);               
                //firmaObject = PDImageXObject.createFromFile(firma, document);
            }
            
            if (chkDC3Logoa.equalsIgnoreCase("true")){
                logoObject = LosslessFactory.createFromImage(document, logo);               
                //logoObject = PDImageXObject.createFromFile(logo, document);
            }
        }catch(Exception ex){
            JOptionPane.showMessageDialog(null, "Error al crear objetos de logo o firma \nfile: " + String.valueOf(logoObject) + "\n" + String.valueOf(firmaObject) + "\n" + ex.toString());
        }

        
        PDPage page = (PDPage)document.getDocumentCatalog().getPages().get(0);         
        PDFont helvetica = PDType1Font.HELVETICA_BOLD;
        PDFont helvetica_normal = PDType1Font.HELVETICA;
        PDPageContentStream contentStream = new PDPageContentStream(document, page, true, true);
        
        
        contentStream.beginText();
        contentStream.setFont( helvetica_normal, 13 );
        contentStream.setNonStrokingColor(Color.BLACK);
        
        contentStream.setTextMatrix(new Matrix(1,0,0,1,30,602 ));           
        contentStream.showText(p.apellidos + " " + p.nombre);
        
        
        char[] curpArray = p.curp.toCharArray();
        float[] xPosition = {32, 49, 63, 77, 92, (float)106.5, 120, (float)134.5, 149, 163, 177, (float)191.5, 205, 223, 241, 255, 269, 286};
        
        float charWidth;
        
        for (int i = 0; i <= curpArray.length - 1; i++) {
            charWidth = helvetica_normal.getStringWidth(String.valueOf(curpArray[i])) /1000 * 13;
            //System.out.println("Char " + i + " width: " + charWidth);
            
            float x = xPosition[i] - (charWidth/2);
            contentStream.setTextMatrix(new Matrix(1,0,0,1, x ,572  ));           
            contentStream.showText(String.valueOf(curpArray[i]));
        }     
  
        contentStream.setTextMatrix(new Matrix(1,0,0,1,30,548 ));           
        contentStream.showText(p.area_puesto);  
        
        contentStream.setTextMatrix(new Matrix(1,0,0,1,30,486 ));           
        if (c.walmart){
            contentStream.showText("OPERADORA WALMART S DE RL DE CV");
        }else{
            contentStream.showText(c.nombre_empresa);
        }
        
        char[] rfc_emprea_array;
                                            
        if (c.walmart){
            rfc_emprea_array = "OWM011023AWA".toCharArray();
        }else{
            rfc_emprea_array = c.rfc_empresa.toCharArray();
        }
        
        try{
        int j = 0;
        for (int i = 0; i <= rfc_emprea_array.length - 1; i++) {
            charWidth = helvetica_normal.getStringWidth(String.valueOf(rfc_emprea_array[rfc_emprea_array.length - i - 1])) /1000 * 13;
            //System.out.println("Char " + i + " width: " + charWidth);
            
            if (j == 3 || i == 9){
                j=j+1;
            }
            float x = xPosition[xPosition.length - j - 4] - (charWidth/2);
            contentStream.setTextMatrix(new Matrix(1,0,0,1, x ,458  ));           
            contentStream.showText(String.valueOf(rfc_emprea_array[rfc_emprea_array.length - i - 1]));
           j=j+1;
        } 
        } catch (Exception ex){
            JOptionPane.showMessageDialog(null, "El RFC de la empresa esta mal formado. \n\nex.toString : " + ex.toString());
            return;
        }
        contentStream.setTextMatrix(new Matrix(1,0,0,1,30,405 ));           
        contentStream.showText(c.nombre_curso);
        
        contentStream.setTextMatrix(new Matrix(1,0,0,1,30,381 ));           
        contentStream.showText(c.horas_texto);
        
        //if (c.walmart){
            contentStream.setTextMatrix(new Matrix(1,0,0,1,30,356 ));           
            contentStream.showText("6000 SEGURIDAD");
        //} else{
        //    contentStream.setTextMatrix(new Matrix(1,0,0,1,30,356 ));           
        //    contentStream.showText(p.area_tematica);
        //}
        
        contentStream.setTextMatrix(new Matrix(1,0,0,1,30,332 ));           
        contentStream.showText(c.uCapacitadora);
        
        
        Calendar cal = Calendar.getInstance();
        cal.setTime(c.fecha_inicio);
        String month = String.format("%02d",cal.get(Calendar.MONTH) + 1);
        String year = String.format("%04d",cal.get(Calendar.YEAR));
        String day = String.format("%02d",cal.get(Calendar.DAY_OF_MONTH));
        
        String date = year + month + day;
        
        cal.setTime(c.fecha_termino);
        month = String.format("%02d",cal.get(Calendar.MONTH) + 1);
        year = String.format("%04d",cal.get(Calendar.YEAR));
        day = String.format("%02d",cal.get(Calendar.DAY_OF_MONTH));
        
        date = date + year + month + day;
                
        char[] date_array = date.toCharArray();
        float[] xPos = {256, 272, 288, 304, 322, 343, 365, 386, 428, 447, 467, 486, 507, 528, 549, 570};
        
        for (int i = 0; i <= date_array.length - 1; i++) {
            charWidth = helvetica_normal.getStringWidth(String.valueOf(date_array[i])) /1000 * 13;
            //System.out.println("Char " + i + " width: " + charWidth);
           
            float x = xPos[i] - (charWidth/2);
            contentStream.setTextMatrix(new Matrix(1,0,0,1, x ,381  ));           
            contentStream.showText(String.valueOf(date_array[i]));
            
        } 
        
        charWidth = helvetica_normal.getStringWidth(p.area_tematica) /1000 * 13;
        //System.out.println(p.area_tematica + " bold : " + charWidth);
        if (charWidth >= 255){
            contentStream.setFont( helvetica_normal, 10 );
            charWidth = helvetica_normal.getStringWidth(p.area_tematica) /1000 * 10;
            //System.out.println(p.area_tematica + " : " + charWidth);
            if (charWidth >= 260){
                contentStream.setFont( helvetica_normal, 8 );
                //charWidth = helvetica_normal.getStringWidth(p.area_tematica) /1000 * 8;
                //System.out.println(p.area_tematica + " : " + charWidth);
            }
        }
        
        
        contentStream.setTextMatrix(new Matrix(1,0,0,1,310,572 ));           
        contentStream.showText(p.area_tematica); 
        
        contentStream.endText();
        
        
        //xPosition = 256, 272, 288, 304, 322, 343, 365, 386, 428, 447, 467, 486, 507, 528, 549, 570
        //year: 256, 272, 288, 304
        //month: 322, 343
        //day: 365, 386
        
        //year: 428, 447, 467, 486
        //month: 507, 528
        //day: 549, 570

        
        // 32, 49, 63, 77, 92, 106.5, 120, 134.5, 149, 163, 177, 191.5, 205, 223, 241, 255, 269, 286  
        /*Esto es para medir las casillas del curp
        contentStream.setStrokingColor(Color.BLACK);
        contentStream.moveTo(286,569);
        contentStream.lineTo(286,580);
        contentStream.stroke();
        */
        
        /*
        PDImageXObject logo = PDImageXObject.createFromFile("src/files/logo.png", document);
        PDImageXObject firma = PDImageXObject.createFromFile("src/files/firmasola.png", document);
        
        System.out.println("logoWidth: " + logo.getWidth());
        System.out.println("logoHeight: " + logo.getHeight());
        System.out.println("firmaWidth: " + firma.getWidth());
        System.out.println("firmaHeight: " + firma.getHeight());
        */
        contentStream.addRect(50, 750, 500, 100);
        contentStream.setNonStrokingColor(Color.WHITE);
        contentStream.fill();
        
        contentStream.addRect(50, 224, 150, 10);
        contentStream.setNonStrokingColor(Color.WHITE);
        contentStream.fill();
        
        if (logoObject != null && !logoObject.isEmpty()){
            contentStream.drawImage(logoObject, 430,700,150,75);
        }
        if (firmaObject != null && !firmaObject.isEmpty()){
            contentStream.drawImage(firmaObject, 80,223,110,42);
        
            contentStream.setStrokingColor(Color.BLACK);
            contentStream.setLineWidth((float).8);
            contentStream.moveTo(50,(float)235.6);
            contentStream.lineTo(200,(float)235.6);
            contentStream.stroke();
            
        }
  
        contentStream.beginText();
        contentStream.setFont( helvetica_normal, 8 );
        contentStream.setNonStrokingColor(Color.BLACK);

        charWidth = helvetica_normal.getStringWidth(c.capacitador) /1000 * 8;
        float x = 126 - charWidth/2;
                
        contentStream.setTextMatrix(new Matrix(1,0,0,1,x,228 ));           
        contentStream.showText(c.capacitador); 
        //GIS100219KK8003
        String regUnidad;
        
        if (c.uCapacitadora.equalsIgnoreCase("TSI. Jorge Antonio Razón Gil")){
            regUnidad = "RAGJ610813BIA005";
        } else {
            regUnidad = "GIS100219KK8003";
        }
        
        charWidth = helvetica_normal.getStringWidth(regUnidad) /1000 * 8;
        x = 126 - charWidth/2;
        contentStream.setTextMatrix(new Matrix(1,0,0,1,x,219 ));           
        contentStream.showText(regUnidad); 
        
        contentStream.endText();
        
         // Make sure that the content stream is closed:
        contentStream.close();
        Format formatter = new SimpleDateFormat("ddMMMYYYY", new Locale("es","MX"));
        String formatedDate = formatter.format(c.fecha_inicio);
        
        String abrev = abreviaturas.get(c.nombre_curso);
        
        // Save the results and ensure that the document is properly closed:
        
        document.save(savePath + File.separator + "p_DC3_" + p.curp + "_" + p.determinante + "_" + p.nombre.replaceAll(" ","_") + "_" + p.apellidos.replaceAll(" ","_") + "_" + abrev + "_" + formatedDate + ".pdf");
        document.close();        
        dosc.put(savePath + File.separator + "p_DC3_" + p.curp + "_" + p.determinante + "_" + p.nombre.replaceAll(" ","_") + "_" + p.apellidos.replaceAll(" ","_") + "_" + abrev + "_" + formatedDate + ".pdf", p.determinante);
        
        
    }
    private static void imprimirConstancias(ArrayList <Participante> listaParticipantes, Curso c, ArrayList <Directorio> listaDirectorio, String chkConstFirma, String chkConstLogo, String savePath, Map<String,String> dosc, String instructor, Map<String, String> abreviaturas) throws IOException {
        
        
        ArrayList <Directorio> listaDirecciones = new ArrayList <> ();
        
        ListIterator <Participante> itParticipantes = listaParticipantes.listIterator(); 
        ListIterator <Directorio> itDirectorio; 
        
        while (itParticipantes.hasNext()){
            Participante p = itParticipantes.next();
            itDirectorio = listaDirectorio.listIterator(); 
            while (itDirectorio.hasNext()){
                Directorio d = itDirectorio.next();
                
                if (p.determinante.equalsIgnoreCase(d.determinante) && !listaDirecciones.contains(d)){
                    listaDirecciones.add(d);
                    break;
                }
            }
        }
        
        
        ListIterator <Directorio> itDireccionesConstancia = listaDirecciones.listIterator(); 
        while (itDireccionesConstancia.hasNext()){
            Directorio d = itDireccionesConstancia.next();
            imprimirUnaConstancia(listaParticipantes, c, d, chkConstFirma, chkConstLogo, savePath, dosc, instructor, abreviaturas);
        }
      
    }
    private static void imprimirUnaConstancia(ArrayList <Participante> listaParticipantes, Curso c, Directorio d, String chkConstFirma, String chkConstLogo, String savePath, Map<String,String> dosc, String instructor, Map<String, String> abreviaturas) throws IOException {
        
        String contanciaTemplate = "";
        
        switch (c.nombre_curso){
            case "PREVENCIÓN Y COMBATE DE INCENDIOS I" : contanciaTemplate = "files/certificado_vacio_incendio_basico_nf_nl.pdf";            
            break;
            case "BUSQUEDA Y RESCATE" : contanciaTemplate = "files/certificado_vacio_busq_rescate_nf_nl.pdf";            
            break;
            case "EVACUACIÓN, BUSQUEDA Y RESCATE" : contanciaTemplate = "files/certificado_vacio_evac_busq_resc_nf_nl.pdf";            
            break;
            case "EVACUACIÓN" : contanciaTemplate = "files/certificado_vacio_evacuacion_nf_nl.pdf";            
            break;
            case "PREVENCIÓN Y COMBATE DE INCENDIOS II" : contanciaTemplate = "files/certificado_vacio_incendio_intermedio_nf_nl.pdf";            
            break;
            case "PREVENCIÓN Y COMBATE DE INCENDIOS III" : contanciaTemplate = "files/certificado_vacio_incendio_avanzado_nf_nl.pdf";            
            break;
            case "FORMACION DE BRIGADAS MULTIFUNCIONALES DE EMERGENCIA" : contanciaTemplate = "files/certificado_vacio_multi_nf_nl.pdf";            
            break;
            case "FORMACIÓN DE BRIGADA MULTIFUNCIONAL DE EMERGENCIA" : contanciaTemplate = "files/certificado_vacio_multi_nf_nl.pdf";            
            break;
            case "FORMACION DE BRIGADA MULTIFUNCIONAL DE EMERGENCIAS" : contanciaTemplate = "files/certificado_vacio_multi_nf_nl.pdf";      
            break;
            case "PRIMEROS AUXILIOS" : contanciaTemplate = "files/certificado_vacio_primeros_auxilios_nf_nl.pdf";            
            break;
            default : contanciaTemplate = "files/cerificado_vacio.pdf";
            break;
        }
        
        InputStream file = null;
        try {
            file = cdiscisa.Cdiscisa.class.getClassLoader().getResourceAsStream(contanciaTemplate);
        }
        catch(Exception ex){
            JOptionPane.showMessageDialog(null, "Error al cargar el certificado base. \nconstancia: " + contanciaTemplate +  "\nfile: " + String.valueOf(file) + "\n" + ex.toString());
        }
        PDDocument document = PDDocument.load(file);
        PDPage page1 = (PDPage)document.getDocumentCatalog().getPages().get(0);
        PDPage page2 = (PDPage)document.getDocumentCatalog().getPages().get(1);

        PDPageContentStream contentStream = new PDPageContentStream(document, page1, true, true);
        PDPageContentStream contentStream2 = new PDPageContentStream(document, page2, true, true);
        
        float pageWidth = page1.getMediaBox().getWidth();
        
        BufferedImage logo = null;
        BufferedImage firma = null;
        
        try{
            ClassLoader classLoader = Thread.currentThread().getContextClassLoader();
            
            //logo = new File(classLoader.getResource("files/logo.png").getFile());
            //logo = StreamUtil.stream2file(cdiscisa.Cdiscisa.class.getClassLoader().getResourceAsStream("files/logo.png"));
           logo = ImageIO.read(cdiscisa.Cdiscisa.class.getClassLoader().getResourceAsStream("files/logo.png"));
            
            if(instructor.equalsIgnoreCase("Ing. Jorge Antonio Razón Gutierrez")){
                 firma = ImageIO.read(cdiscisa.Cdiscisa.class.getClassLoader().getResourceAsStream("files/firmaCoco.png"));
                 //firma = new File(cdiscisa.Cdiscisa.class.getClassLoader().getResource("files/firmaCoco.png").getFile());
             } else if (instructor.equalsIgnoreCase("Manuel Anguiano Razón")){
                 firma = ImageIO.read(cdiscisa.Cdiscisa.class.getClassLoader().getResourceAsStream("files/firmaManuel.png"));
                 //firma = new File(cdiscisa.Cdiscisa.class.getClassLoader().getResource("files/firmaManuel.png").getFile());
             } else {
                 firma = ImageIO.read(cdiscisa.Cdiscisa.class.getClassLoader().getResourceAsStream("files/firmaJorge.png"));
                 //firma = new File(cdiscisa.Cdiscisa.class.getClassLoader().getResource("files/firmaJorge.png").getFile());
            }
             
        }catch (Exception ex){
            JOptionPane.showMessageDialog(null, "Error al cargar la imagen del logo o la firma \nfile: " + String.valueOf(logo) + "\n" + String.valueOf(firma) + "\n" + ex.toString());
        }
        
        PDImageXObject firmaObject = null;
        PDImageXObject logoObject = null;
        
        try{
            if (chkConstFirma.equalsIgnoreCase("true")){
                firmaObject = LosslessFactory.createFromImage(document, firma);
                //firmaObject = PDImageXObject.createFromFile(firma, document);
            }
            if (chkConstLogo.equalsIgnoreCase("true")){
                logoObject = LosslessFactory.createFromImage(document, logo);
                //logoObject = PDImageXObject.createFromFile(logo, document);
            }
        }catch(Exception ex){
            JOptionPane.showMessageDialog(null, "Error al crear objetos de logo o firma \nfile: " + String.valueOf(logoObject) + "\n" + String.valueOf(firmaObject) + "\n" + ex.toString());
        }
        
        InputStream isFont1 = null, isFont2 = null;
        try{
        isFont1 = cdiscisa.Cdiscisa.class.getClassLoader().getResourceAsStream("files/Calibri.ttf");       
        isFont2 = cdiscisa.Cdiscisa.class.getClassLoader().getResourceAsStream("files/CalibriBold.ttf");
        } catch(Exception ex) {
            JOptionPane.showMessageDialog(null, "Error al cargar el una fuente \nisFont1: " + String.valueOf(isFont1) + "\nisFont2: " + String.valueOf(isFont2) +  "\n" + ex.toString());
        }
        
        PDFont calibri = PDType0Font.load(document, isFont1);
        PDFont calibriBold = PDType0Font.load(document, isFont2);
        
        
        contentStream.beginText();
        DateFormat df = DateFormat.getDateInstance(DateFormat.LONG, new Locale("es","MX"));
                
        contentStream.setFont( calibri, 9 );
        contentStream.setNonStrokingColor(Color.BLACK);
        
        contentStream.setTextMatrix(new Matrix(1,0,0,1, 465, 656)); 
        contentStream.showText(df.format(new Date()));
                
                
        contentStream.setFont( calibriBold, 11 );
        
        
        contentStream.setTextMatrix(new Matrix(1,0,0,1, 135,585)); 
        contentStream.showText(c.razon_social);
        
        contentStream.setTextMatrix(new Matrix(1,0,0,1,120, (float)572.5)); 
        if (c.walmart){
            contentStream.showText(d.determinante + " " + d.unidad);            
        } else if (d.sucursal.isEmpty() || d.sucursal.equalsIgnoreCase("")){
            contentStream.endText();
            contentStream.addRect(30, 572, 150, 10);
            contentStream.setNonStrokingColor(Color.WHITE);
            contentStream.fill();
            contentStream.beginText();
            contentStream.setNonStrokingColor(Color.BLACK);
        } else{
            contentStream.showText(d.sucursal);                                   
        }   
        
        if (!c.walmart){
            contentStream.setTextMatrix(new Matrix(1,0,0,1,120, 527)); 
            contentStream.showText(d.RFC);
            contentStream.setFont(calibri, 11);
            contentStream.setTextMatrix(new Matrix(1,0,0,1,72, 527)); 
            contentStream.showText("RFC: ");
            contentStream.setFont( calibriBold, 11 ); 
        }
        /*
        if (c.walmart){
            contentStream.showText(p1.determinante + " " + d.unidad);
        } else if (d.sucursal.isEmpty() || d.sucursal.equalsIgnoreCase("")){
            contentStream.endText();
            contentStream.addRect(100, 590, 400, 20);
            contentStream.setNonStrokingColor(Color.WHITE);
            contentStream.fill();
            contentStream.beginText();
            contentStream.setNonStrokingColor(Color.BLACK);
        } else{
            contentStream.showText(d.sucursal);
        }
        
        */
        
        float charWidth = calibriBold.getStringWidth(d.direccion) /1000 * 11;
        //System.out.println(charWidth + " " + d.direccion.length() + " " +  d.direccion);

        if (charWidth <= 400){
            contentStream.setTextMatrix(new Matrix(1,0,0,1,120, 549));           
            contentStream.showText(d.direccion);
        } else {
            contentStream.setTextMatrix(new Matrix(1,0,0,1,120, 552));           
            contentStream.showText(d.direccion.substring(0, d.direccion.indexOf(" ", 82)));
            contentStream.setTextMatrix(new Matrix(1,0,0,1,120, 541));           
            contentStream.showText(d.direccion.substring(d.direccion.indexOf(" ", 82) + 1, d.direccion.length()));
        }
        
        charWidth = calibriBold.getStringWidth(c.nombre_curso) /1000 * 11;
        
        
        float xPosition = (pageWidth - charWidth)/2;
                
        contentStream.setTextMatrix(new Matrix(1,0,0,1,xPosition, 490));           
        contentStream.showText(c.nombre_curso);
        
        contentStream.setTextMatrix(new Matrix(1,0,0,1,160, 465));           
        contentStream.showText(c.fecha_certificado);
        
        contentStream.setTextMatrix(new Matrix(1,0,0,1,160, (float)450.5));           
        contentStream.showText(c.horas_texto);
        
        ListIterator <Participante> it = listaParticipantes.listIterator();
        
        float y = 0;
        while(it.hasNext()){
            Participante p = it.next();
            
            if (p.determinante.equalsIgnoreCase(d.determinante)){
                contentStream.setFont( calibri, 11 );
                
                charWidth = calibri.getStringWidth(p.nombre + " " + p.apellidos) /1000 * 11;                
                //System.out.println(charWidth + " " + p.nombre + " " + p.apellidos);
                
                if (charWidth > 165){
                    contentStream.setFont( calibri, 9 );
                    contentStream.setTextMatrix(new Matrix(1,0,0,1,135,(float)376.5 - y));           
                    contentStream.showText(p.nombre + " " + p.apellidos);
                    
                    contentStream.setFont( calibri, 11 );
                    
                } else{
                    contentStream.setTextMatrix(new Matrix(1,0,0,1,135, (float)376.5 - y));           
                    contentStream.showText(p.nombre + " " + p.apellidos);
                }
                
                charWidth = calibri.getStringWidth(p.area_puesto) /1000 * 11;
                //System.out.println(charWidth + " " + p.area_puesto);
                
                if (charWidth > 112){
                    contentStream.setFont( calibri, 9 );
                    contentStream.setTextMatrix(new Matrix(1,0,0,1,360, (float)376.5 - y));           
                    contentStream.showText(p.area_puesto);
                
                } else {
                    contentStream.setTextMatrix(new Matrix(1,0,0,1,360, (float)376.5 - y));           
                    contentStream.showText(p.area_puesto);
                }
                
                y = y + (float)12.7;
                 
            }
        }
        
        contentStream.endText();
        float nameWidth;
        
        if (logoObject != null && !logoObject.isEmpty()){
            contentStream.drawImage(logoObject, 30,700,156,78);
        }
        if (firmaObject != null && !firmaObject.isEmpty()){
            xPosition = (pageWidth - 100)/2;
            contentStream.drawImage(firmaObject, xPosition,55,110,42);

            contentStream.beginText();
            contentStream.setFont(calibriBold, 10);
            contentStream.setNonStrokingColor(Color.BLACK);
            
            nameWidth = calibriBold.getStringWidth(instructor) /1000 * 10;            
            xPosition = (pageWidth - nameWidth)/2 + 9;            
            contentStream.setTextMatrix(new Matrix(1,0,0,1,xPosition, 55 ));           
            contentStream.showText(instructor);
            
            if (instructor.equalsIgnoreCase("Manuel Anguiano Razón")){
                nameWidth = calibriBold.getStringWidth(c.registro_manuel) /1000 * 10;            
                xPosition = (pageWidth - nameWidth)/2 + 9; 
                contentStream.setTextMatrix(new Matrix(1,0,0,1,xPosition, 40 ));           
                contentStream.showText(c.registro_manuel);
            } else if (instructor.equalsIgnoreCase("Ing. Jorge Antonio Razón Gutierrez")){
                nameWidth = calibriBold.getStringWidth(c.registro_coco) /1000 * 10;            
                xPosition = (pageWidth - nameWidth)/2 + 9; 
                contentStream.setTextMatrix(new Matrix(1,0,0,1,xPosition, 40 ));           
                contentStream.showText(c.registro_coco);
            } else {
                nameWidth = calibriBold.getStringWidth(c.registro_jorge) /1000 * 10;            
                xPosition = (pageWidth - nameWidth)/2 + 9; 
                contentStream.setTextMatrix(new Matrix(1,0,0,1,xPosition, 40 ));           
                contentStream.showText(c.registro_jorge);
            }

            contentStream.endText();
        }
        
        // Make sure that the content stream is closed:
        contentStream.close();
        
        
        contentStream2.beginText();
        
        contentStream2.setFont( calibri, 11 );
        contentStream2.setNonStrokingColor(Color.BLACK);
        
        contentStream2.setTextMatrix(new Matrix(1,0,0,1, 465, 656)); 
        contentStream2.showText(df.format(new Date()));
                
                
        contentStream2.setFont( calibriBold, 11 );
        
        contentStream2.setTextMatrix(new Matrix(1,0,0,1, 135,585)); 
        contentStream2.showText(c.razon_social);
        
        contentStream2.setTextMatrix(new Matrix(1,0,0,1,120, (float)572.5));           
        if (c.walmart){
            contentStream2.showText(d.determinante + " " + d.unidad);            
        } else{
            contentStream2.showText(d.sucursal);
            contentStream2.setTextMatrix(new Matrix(1,0,0,1,120, 527)); 
            contentStream2.showText(d.RFC);
            contentStream2.setFont(calibri, 11);
            contentStream2.setTextMatrix(new Matrix(1,0,0,1,73, 527)); 
            contentStream2.showText("RFC: ");      
            contentStream2.setFont( calibriBold, 11 );
        }        
        
        charWidth = calibriBold.getStringWidth(d.direccion) /1000 * 11;
        //System.out.println(charWidth + " " + d.direccion.length() + " " +  d.direccion);

        if (charWidth <= 400){
            contentStream2.setTextMatrix(new Matrix(1,0,0,1,120, 549));           
            contentStream2.showText(d.direccion);
        } else {
            contentStream2.setTextMatrix(new Matrix(1,0,0,1,120, 555));           
            contentStream2.showText(d.direccion.substring(0, d.direccion.indexOf(" ", 82)));
            contentStream2.setTextMatrix(new Matrix(1,0,0,1,120, 540));           
            contentStream2.showText(d.direccion.substring(d.direccion.indexOf(" ", 82) + 1, d.direccion.length()));
        }
        
        charWidth = calibriBold.getStringWidth(c.nombre_curso) /1000 * 11;
        
        pageWidth = page2.getMediaBox().getWidth();
        xPosition = (pageWidth - charWidth)/2;
                
        contentStream2.setTextMatrix(new Matrix(1,0,0,1,xPosition, 490));           
        contentStream2.showText(c.nombre_curso);
        
        contentStream2.setTextMatrix(new Matrix(1,0,0,1,160, 465));           
        contentStream2.showText(c.fecha_certificado);
        
        contentStream2.setTextMatrix(new Matrix(1,0,0,1,160, (float)450.5));           
        contentStream2.showText(c.horas_texto);
               
        contentStream2.endText();
        
        if (logoObject != null && !logoObject.isEmpty()){
            contentStream2.drawImage(logoObject, 30,700,156,78);
        }
        if (firmaObject != null && !firmaObject.isEmpty()){
            xPosition = (pageWidth - 100)/2;
            contentStream2.drawImage(firmaObject, xPosition,55,110,42);

            contentStream2.beginText();
            contentStream2.setFont(calibriBold, 10);
            contentStream2.setNonStrokingColor(Color.BLACK);
            
            nameWidth = calibriBold.getStringWidth(instructor) /1000 * 10;            
            xPosition = (pageWidth - nameWidth)/2 + 9;            
            contentStream2.setTextMatrix(new Matrix(1,0,0,1,xPosition, 55 ));           
            contentStream2.showText(instructor);
            
            if (instructor.equalsIgnoreCase("Manuel Anguiano Razón")){
                nameWidth = calibriBold.getStringWidth(c.registro_manuel) /1000 * 10;            
                xPosition = (pageWidth - nameWidth)/2 + 9; 
                contentStream2.setTextMatrix(new Matrix(1,0,0,1,xPosition, 40 ));           
                contentStream2.showText(c.registro_manuel);
            } else if (instructor.equalsIgnoreCase("Ing. Jorge Antonio Razón Gutierrez")){
                nameWidth = calibriBold.getStringWidth(c.registro_coco) /1000 * 10;            
                xPosition = (pageWidth - nameWidth)/2 + 9; 
                contentStream2.setTextMatrix(new Matrix(1,0,0,1,xPosition, 40 ));           
                contentStream2.showText(c.registro_coco);
            } else {
                nameWidth = calibriBold.getStringWidth(c.registro_jorge) /1000 * 10;            
                xPosition = (pageWidth - nameWidth)/2 + 9; 
                contentStream2.setTextMatrix(new Matrix(1,0,0,1,xPosition, 40 ));           
                contentStream2.showText(c.registro_jorge);
            }
            contentStream2.endText();
        }
            
            
        contentStream2.close();
        
        //"Capacitacion_BAE_Centro de Huinala_2631_MULTI_19ago2015"

        //Capacitación + formato tienda + nombre sucursal + numero sucursal + nombre curso + ddmmaaaa
       
        Format formatter = new SimpleDateFormat("ddMMMYYYY", new Locale("es","MX"));
        String formatedDate = formatter.format(c.fecha_inicio);
        
        String abrev = abreviaturas.get(c.nombre_curso);
        
        // Save the results and ensure that the document is properly closed:
        if(c.walmart){
            document.save(savePath + File.separator + "Certificado_" + d.formato + "_" + d.unidad + "_" + d.determinante + "_" + abrev + "_" + formatedDate + ".pdf");
            document.close();
            dosc.put(savePath + File.separator + "Certificado_" + d.formato + "_" + d.unidad + "_" + d.determinante + "_" + abrev + "_" + formatedDate + ".pdf", d.determinante);
        }else{
            document.save(savePath + File.separator + "Certificado_" + d.determinante + "_" + abrev + "_" + formatedDate + ".pdf");
            document.close();
            dosc.put(savePath + File.separator + "Certificado_" + d.determinante + "_" + abrev + "_" + formatedDate + ".pdf", d.determinante);
        }
        
    }
    private static void imprimirDiplomas_main(ArrayList <Participante> listaParticipantes, Curso c, ArrayList <Directorio> listaDirectorio, String chkDipFirma, String chkDipLogo, String savePath, Map<String,String> dosc, String instructor, Map<String, String> abreviaturas) throws IOException{
        
       
        
        ArrayList <Directorio> listaDirecciones = new ArrayList <> ();
        ArrayList <Participante> participantesSucursal = new ArrayList <> ();
        
        ListIterator <Participante> itParticipantes = listaParticipantes.listIterator(); 
        ListIterator <Directorio> itDirectorio; 
         
        
        while (itParticipantes.hasNext()){
            Participante p = itParticipantes.next();
            itDirectorio = listaDirectorio.listIterator(); 
            while (itDirectorio.hasNext()){
                Directorio d = itDirectorio.next();
                
                if (p.determinante.equalsIgnoreCase(d.determinante) && !listaDirecciones.contains(d)){
                    listaDirecciones.add(d);
                    break;
                }
            }
        }
        
        ListIterator <Directorio> itDireccionesConstancia = listaDirecciones.listIterator(); 
        while (itDireccionesConstancia.hasNext()){
            Directorio d = itDireccionesConstancia.next();
            
            itParticipantes = listaParticipantes.listIterator(); 
            while(itParticipantes.hasNext()){
                Participante p1 = itParticipantes.next();
                if (d.determinante.equalsIgnoreCase(p1.determinante)){
                    participantesSucursal.add(p1);
                }
            }
            imprimirDiplomas(participantesSucursal, c, d, chkDipFirma, chkDipLogo, savePath, dosc, instructor, abreviaturas);
            participantesSucursal.clear();
        }
    }
    private static void mergeFiles(Map<String,String> dosc, ArrayList <Participante> listaParticipantes, ArrayList <Directorio> listaDirectorio, String registroPDF, String lista_auto_pdf ) throws FileNotFoundException, IOException{
        
        ArrayList <Directorio> listaDirecciones = new ArrayList <> ();
        
        ListIterator <Participante> itParticipantes = listaParticipantes.listIterator(); 
        ListIterator <Directorio> itDirectorio; 
        
        while (itParticipantes.hasNext()){
            Participante p = itParticipantes.next();
            itDirectorio = listaDirectorio.listIterator(); 
            while (itDirectorio.hasNext()){
                Directorio d = itDirectorio.next();
                
                if (p.determinante.equalsIgnoreCase(d.determinante) && !listaDirecciones.contains(d)){
                    listaDirecciones.add(d);
                    break;
                }
            }
        }
        
        
        ListIterator <Directorio> itDireccionesConstancia = listaDirecciones.listIterator(); 
        while (itDireccionesConstancia.hasNext()){
            
            PDFMergerUtility merge = new PDFMergerUtility();
            
            Directorio d = itDireccionesConstancia.next();
            String x,name;
            name = "";
            
            SortedSet<String> keys = new TreeSet<String>(dosc.keySet());
            
            for (String key : keys) { 
            String value = dosc.get(key);
                if (value.equalsIgnoreCase(d.determinante)) {
                   
                    if (key.contains("Diplomas_")){
                        name = key.replaceFirst("Diplomas_", "Capacitacion_");
                    }
                merge.addSource(new File(key));
                }
                
            }
            /*
            for (Map.Entry<String, String> entry : dosc.entrySet()) {
                if (entry.getValue().equals(d.determinante)) {
                    x = entry.getKey();
                    if (x.contains("Diplomas_")){
                        name = x.replaceFirst("Diplomas_", "Capacitacion_");
                    }
                merge.addSource(new File(x));
                }
            }
            */
            
            //InputStream newReg = registroPDF;
            //InputStream newListaAuto = lista_auto_pdf;
            
            
            
            try{
                merge.addSource(new File(lista_auto_pdf));
                
                if(registroPDF.equalsIgnoreCase("files/RegistroPCJorge Razon2016.pdf")){
                    merge.addSource(cdiscisa.Cdiscisa.class.getClassLoader().getResourceAsStream(registroPDF));
                } else{
                    merge.addSource(new File(registroPDF));
                }
                
                merge.setDestinationFileName(name);
                merge.mergeDocuments(MemoryUsageSetting.setupMainMemoryOnly());
            } catch (Exception ex){
                JOptionPane.showMessageDialog(null, "registroPDF: " + registroPDF + "\nlista_auto_pdf: " + lista_auto_pdf + "\n" + ex.getMessage());
            }
            
            //newReg = null;
            //newListaAuto = null;
            merge = null;
            
        }
        
        
        
        
        
  
        //"Capacitacion_BAE_Centro de Huinala_2631_MULTI_19ago2015"

        //Capacitación + formato tienda + nombre sucursal + numero sucursal + nombre curso + ddmmaaaa
    
        // Save the results and ensure that the document is properly closed:
        //document.save(savePath + File.separator + "Certificado_" + d.formato + "_" + d.unidad + "_" + d.determinante + "_MULTI_" + formatedDate + ".pdf");
        
    }
}
