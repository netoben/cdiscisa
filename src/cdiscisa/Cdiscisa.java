/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package cdiscisa;

//import static com.sun.org.apache.bcel.internal.util.SecuritySupport.getResourceAsStream;
import java.awt.Color;


import java.io.File;
import java.io.FileInputStream;

import java.io.FileNotFoundException;
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
import javax.swing.JOptionPane;
import org.apache.pdfbox.cos.COSDictionary;
import org.apache.pdfbox.io.MemoryUsageSetting;
import org.apache.pdfbox.multipdf.PDFMergerUtility;

/**
 *
 * @author Ernesto Armendáriz Bernal.
 */

class Directorio {
String determinante,formato,unidad,estado,municipio,direccion;

public Directorio(){
    this.determinante = "";
    this.direccion = "";
    this.estado = "";
    this.formato = "";
    this.municipio = "";
    this.unidad = "";
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
    String razon_social,rfc_empresa, fecha_certificado, capacitador, uCapacitadora;
    int horas;
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
        
        String registroPDF = "";
        InputStream lista_auto_pdf = null;
        
        String nameRegistro = "";
        
        if (args[15].isEmpty()){
            nameRegistro = "files/registro_jorge_2015.pdf";
        } else {        
            nameRegistro = args[15];
        }
        /*
        try{
            registroPDF = 
           // registroPDF = cdiscisa.Cdiscisa.class.getClassLoader().getResourceAsStream(nameRegistro);
        } catch (Exception ex){
            JOptionPane.showMessageDialog(null, "Error al cargar el registro default Jorge Razon 2015. \nnameRegistro: " + nameRegistro +  "\nfile: " + String.valueOf(registroPDF) +  "\n" + ex.toString());
            return;
        }
        */
        if (args[16].isEmpty()){
            JOptionPane.showMessageDialog(null, "Es necesario proporcionar la lista autógrafa de este curso.");
            return;
        } else {
            //try{
            /*    lista_auto_pdf = new FileInputStream(new File(args[16]));
            } catch (Exception ex){
                JOptionPane.showMessageDialog(null, "Error al cargar la lista autógrafa de este curso. \nargs[15]: " + args[15] +  "\nlista_auto_pdf: " + String.valueOf(lista_auto_pdf) +  "\n" + ex.toString());
            }*/
        }
        
        ArrayList <Directorio> listaDirectorio = null;
        ArrayList <Participante> listaParticipantes = null;
        Curso c = null;
        
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
        
        Map<String,String> dosc = new HashMap<String,String>();
        
        if (args[5].equalsIgnoreCase("true")){            
            imprimirDiplomas_main(listaParticipantes, c, listaDirectorio, args[6], args[7], args[1], dosc);            
        }
        if (args[8].equalsIgnoreCase("true")){            
            imprimirConstancias(listaParticipantes, c, listaDirectorio, args[9], args[10], args[1], dosc);
        }
        if (args[11].equalsIgnoreCase("true")){            
            imprimirDC3(listaParticipantes,c, args[12], args[13], args[1], dosc);
        }
        if (args[14].equalsIgnoreCase("true")){            
            mergeFiles(dosc,listaParticipantes,listaDirectorio, nameRegistro, args[16] );
        }
        
        JOptionPane.showMessageDialog(null, "Los documentos se han generado exitosamente");

    }
    
    private static Curso llenarCurso (Workbook wbLista, String unidadCapacitadora, String instructor) throws Exception{
        
        Sheet wbListaSheet = wbLista.getSheetAt(0);
        
        Curso c = new Curso();
        
        if (!(wbListaSheet.getRow(2) == null || wbListaSheet.getRow(2).getCell(4) == null || wbListaSheet.getRow(2).getCell(4).getStringCellValue().isEmpty()) ){
            c.nombre_empresa = wbListaSheet.getRow(2).getCell(4).getStringCellValue();
        }
        
        if (!(wbListaSheet.getRow(4) == null || wbListaSheet.getRow(4).getCell(4) == null || wbListaSheet.getRow(4).getCell(4).getStringCellValue().isEmpty()) ) {
            c.nombre_curso = wbListaSheet.getRow(4).getCell(4).getStringCellValue();
        }
        
        if (!(wbListaSheet.getRow(6) == null || wbListaSheet.getRow(6).getCell(4) == null || wbListaSheet.getRow(6).getCell(4).getStringCellValue().isEmpty()) ){
            c.nombre_instructor = wbListaSheet.getRow(6).getCell(4).getStringCellValue();
        }
        
        if (!(wbListaSheet.getRow(8) == null || wbListaSheet.getRow(8).getCell(4) == null || wbListaSheet.getRow(8).getCell(4).getStringCellValue().isEmpty())  ){
            c.horas_texto = wbListaSheet.getRow(8).getCell(4).getStringCellValue();
        }
        
        if (!(wbListaSheet.getRow(11) == null || wbListaSheet.getRow(11).getCell(4) == null || wbListaSheet.getRow(11).getCell(4).getStringCellValue().isEmpty())  ){
            c.razon_social = wbListaSheet.getRow(11).getCell(4).getStringCellValue();            
        }
        
        if (!(wbListaSheet.getRow(13) == null || wbListaSheet.getRow(13).getCell(4) == null || wbListaSheet.getRow(13).getCell(4).getStringCellValue().isEmpty())  ){
            c.rfc_empresa = wbListaSheet.getRow(13).getCell(4).getStringCellValue();            
        } else {        
            JOptionPane.showMessageDialog(null, "El RFC de la empresa no puede estar vacio"); 
            throw new netoCustomException("Error al leer los datos del curso");
        }
        
        if (!(wbListaSheet.getRow(15) == null || wbListaSheet.getRow(15).getCell(4) == null || wbListaSheet.getRow(15).getCell(4).getStringCellValue().isEmpty())  ){
            c.fecha_certificado = wbListaSheet.getRow(15).getCell(4).getStringCellValue();            
        }
        
        if (!(wbListaSheet.getRow(17) == null || wbListaSheet.getRow(17).getCell(4) == null || wbListaSheet.getRow(17).getCell(4).getStringCellValue().isEmpty())  ){
            c.fecha_texto_diploma = wbListaSheet.getRow(17).getCell(4).getStringCellValue();            
        }
        
        if (!unidadCapacitadora.isEmpty()) {
            c.uCapacitadora = unidadCapacitadora;
        }
        
        if (!instructor.isEmpty()){
            c.capacitador = instructor;
        }
        
        Calendar cal = Calendar.getInstance();
        
        if (!(wbListaSheet.getRow(4) == null || wbListaSheet.getRow(4).getCell(6) == null)){
            cal.set(Calendar.DAY_OF_MONTH, (int)wbListaSheet.getRow(4).getCell(6).getNumericCellValue());
         }
        if (!(wbListaSheet.getRow(6) == null || wbListaSheet.getRow(6).getCell(6) == null)){
            cal.set(Calendar.MONTH, Integer.parseInt(wbListaSheet.getRow(6).getCell(6).getStringCellValue())-1);
        }
        if (!(wbListaSheet.getRow(8) == null || wbListaSheet.getRow(8).getCell(6) == null)){
            cal.set(Calendar.YEAR, (int)wbListaSheet.getRow(8).getCell(6).getNumericCellValue());
        }
        
        c.fecha_inicio  = cal.getTime();
        
        if (!(wbListaSheet.getRow(4) == null || wbListaSheet.getRow(4).getCell(7) == null)){
            cal.set(Calendar.DAY_OF_MONTH, (int)wbListaSheet.getRow(4).getCell(7).getNumericCellValue());
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
            
            if (row.getCell(2) != null){                    
                try{
                    p.determinante = row.getCell(2).getStringCellValue().trim();
                }catch(Exception ex){
                    JOptionPane.showMessageDialog(null, "Error leyendo determinante del archivo Excel de Lista de participantes  ");
                }
            }
            
            if (row.getCell(3) != null){   
                try{
                p.sucursal = row.getCell(3).getStringCellValue().trim();
                }catch(Exception ex){
                    JOptionPane.showMessageDialog(null, "Error leyendo la sucursal del archivo Excel de Lista de participantes  ");
                }
            }
            
            if (row.getCell(4) != null){   
                try{
                p.nombre = row.getCell(4).getStringCellValue().trim();
                }catch(Exception ex){
                    JOptionPane.showMessageDialog(null, "Error leyendo la columna Nombre del archivo Excel de Lista de participantes  ");
                }
            }
            
            if (row.getCell(5) != null){ 
                try{
                p.apellidos = row.getCell(5).getStringCellValue().trim();
                }catch(Exception ex){
                    JOptionPane.showMessageDialog(null, "Error leyendo la columna Apellidos del archivo Excel de Lista de participantes  ");
                }
            }
            
            if (row.getCell(6) != null){ 
                try{
                p.curp = row.getCell(6).getStringCellValue().trim();
                }catch(Exception ex){
                    JOptionPane.showMessageDialog(null, "Error leyendo la columna CURP del archivo Excel de Lista de participantes  ");
                }
            }
            
            if (row.getCell(7) != null){ 
                try{
                p.area_puesto = row.getCell(7).getStringCellValue().trim();
                }catch(Exception ex){
                    JOptionPane.showMessageDialog(null, "Error leyendo la columna Area Puesto del archivo Excel de Lista de participantes  ");
                }
            }
            
            if (row.getCell(8) != null){ 
                try{
                p.area_tematica = row.getCell(8).getStringCellValue().trim();
                }catch(Exception ex){
                    JOptionPane.showMessageDialog(null, "Error leyendo la columna Area Tematica del archivo Excel de Lista de participantes  ");
                }
            }
            
            p.aprovado = false;
            if (row.getCell(9) != null && row.getCell(9).getStringCellValue().equalsIgnoreCase("Aprobado")){                    
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
    private static ArrayList <Directorio> llenarDirectorio (Workbook wbDirectorio) throws Exception{
        ArrayList <Directorio> listaDirectorio = new ArrayList <>();
        Sheet wbListaSheet = wbDirectorio.getSheetAt(0);
        Iterator<Row> rowIterator = wbListaSheet.iterator();
        
        if (rowIterator.hasNext()){
            rowIterator.next();
        }
        
        Row row = null; 
        Directorio d = null;
        
        while(rowIterator.hasNext()){
        
            row = rowIterator.next();
            
            if (row.getCell(0) == null || row.getCell(0).toString().isEmpty() )
            {break;}
            
            d = new Directorio();
            
            if (row.getCell(0) != null){
                row.getCell(0).setCellType(Cell.CELL_TYPE_STRING);
                d.determinante = row.getCell(0).getStringCellValue().trim();
            }
            
            if (row.getCell(1) != null){  
                int x = row.getCell(1).getCellType();
                d.formato= row.getCell(1).getStringCellValue().trim();
            }
            
            if (row.getCell(2) != null){                     
                d.unidad = row.getCell(2).getStringCellValue().trim();
            }
            if (row.getCell(3) != null){                     
                d.estado = row.getCell(3).getStringCellValue().trim();
            }
            if (row.getCell(4) != null){                     
                d.municipio = row.getCell(4).getStringCellValue().trim();
            }
            if (row.getCell(4) != null){                     
                d.direccion = row.getCell(5).getStringCellValue().trim();
            }
            
            listaDirectorio.add(d);
            
        }
        
        return listaDirectorio;
    }
    private static void imprimirDiplomas(ArrayList <Participante> listaParticipantes, Curso c, Directorio d, String chkDipFirma, String chkDipLogo, String savePath, Map<String,String> dosc) throws IOException{
          
        ListIterator <Participante> it = listaParticipantes.listIterator();
        Participante p1 = null;
        Participante p2 = null;
        
        // Create a document and add a page to it
        PDDocument document = new PDDocument();
        PDDocument documentSingle;
        
        InputStream file = null;
        String nameSingle = "";
               
        if (chkDipFirma.equalsIgnoreCase("true") && chkDipLogo.equalsIgnoreCase("true")){
            nameSingle = "files/diploma_simple_vacio.pdf";            
        } else if(chkDipFirma.equalsIgnoreCase("true") ){
            nameSingle = "files/diploma_simple_vacio.pdf"; 
        } else if(chkDipLogo.equalsIgnoreCase("true")){
            nameSingle = "files/diploma_simple_vacio.pdf"; 
        } else{
            nameSingle = "files/diploma_simple_vacio.pdf"; 
        }       
        
        try{
            file = cdiscisa.Cdiscisa.class.getClassLoader().getResourceAsStream(nameSingle);
        } catch (Exception ex){
            JOptionPane.showMessageDialog(null, "Error al cargar el el diploma single base. \ndiploma: " + nameSingle +  "\nfile: " + String.valueOf(file) + "\n" + ex.toString());
        }
                
        documentSingle = PDDocument.load(file);
        
        PDPage pageSingle = (PDPage)documentSingle.getDocumentCatalog().getPages().get(0);  
        COSDictionary pageDictSingle = pageSingle.getCOSObject();
        COSDictionary newPageSingleDict = new COSDictionary(pageDictSingle);
        PDPage templatePageSingle = new PDPage(newPageSingleDict);
        
        // Create a document and add a page to it
        
        PDDocument documentDoble;
        
        InputStream file2 = null;
        String nameDoble = "";
        
        if (chkDipFirma.equalsIgnoreCase("true") && chkDipLogo.equalsIgnoreCase("true")){
            nameDoble = "files/diploma_doble_vacio.pdf";
        } else if(chkDipFirma.equalsIgnoreCase("true") ){
            nameDoble = "files/diploma_doble_vacio.pdf";
        } else if(chkDipLogo.equalsIgnoreCase("true")){
            nameDoble = "files/diploma_doble_vacio.pdf";
        } else{
            nameDoble = "files/diploma_doble_vacio.pdf";
        }
        
        try{
            file2 = cdiscisa.Cdiscisa.class.getClassLoader().getResourceAsStream(nameDoble);
        } catch (Exception ex){
            JOptionPane.showMessageDialog(null, "Error al cargar el el diploma doble base. \ndiploma: " + nameDoble +  "\nfile: " + String.valueOf(file2) + "\n" + ex.toString());
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
                
                imprimirDiplomaDoble(p1,p2,c,d,document,templatePageDoble, calibri,calibriBold,pristina);
            } 
        } else {
            if (listaParticipantes.size() > 1){
                //Lista es impar y contiene mas de 2 participantes.
                while (it.hasNext()){
                    
                    if (it.nextIndex() == listaParticipantes.size() - 1){
                        p1 = it.next();
                        imprimirDiplomaArriba(p1,c,d,document,templatePageSingle,calibri,calibriBold,pristina);
                        break;
                    }
                    
                    p1 = it.next();
                    p2 = it.next();
                
                    imprimirDiplomaDoble(p1,p2,c,d,document,templatePageDoble,calibri,calibriBold,pristina);
                }          
            } else if (listaParticipantes.size() == 1){
                p1 = it.next();
                imprimirDiplomaArriba(p1,c,d,document,templatePageSingle,calibri,calibriBold,pristina);
            }  
            
        }
        
        Format formatter = new SimpleDateFormat("ddMMMYYYY");
        String formatedDate = formatter.format(c.fecha_inicio);
    
        
        document.save(savePath + File.separator + "Diplomas_" + d.formato + "_" + d.unidad + "_" + d.determinante + "_MULTI_" + formatedDate + ".pdf");
        document.close();
        
        dosc.put(savePath + File.separator + "Diplomas_" + d.formato + "_" + d.unidad + "_" + d.determinante + "_MULTI_" + formatedDate + ".pdf",d.determinante);
    }
    private static void imprimirDiplomaArriba(Participante p1, Curso c, Directorio d, PDDocument document, PDPage page, PDFont calibri, PDFont calibriBold, PDFont pristina) throws IOException{
        
        /*ListIterator <Directorio> it = listaDirectorio.listIterator();
        
        Directorio d = null;
        
         while (it.hasNext()){
              d = it.next();
             if (d.determinante.equalsIgnoreCase(p1.determinante)){
                break;
             }
         }*/
        
        COSDictionary pageDict = page.getCOSObject();
        COSDictionary newPageDict = new COSDictionary(pageDict);

        PDPage newPage = new PDPage(newPageDict);
        document.addPage(newPage);

        // Start a new content stream which will "hold" the to be created content
        PDPageContentStream contentStream = new PDPageContentStream(document, newPage, true, true);

        float pageWidth = newPage.getMediaBox().getWidth();
        float pageHeight = newPage.getMediaBox().getHeight();
        
        //System.out.println("pageWidth: " + pageWidth + "\npageHeight: " + pageHeight);

        // Print Name
        contentStream.beginText();
        contentStream.setFont( pristina, 28 );
        contentStream.setNonStrokingColor(0,112,192);
        
        float nameWidth = pristina.getStringWidth(p1.nombre + " " + p1.apellidos) /1000 * 28;
        
        //System.out.println("nameWidth: " + nameWidth);
        
        float xPosition = (pageWidth - nameWidth)/2 + 15;
        
        contentStream.setTextMatrix(new Matrix(1,0,0,1,xPosition,622 ));           
        contentStream.showText(p1.nombre + " " + p1.apellidos);
        
        contentStream.setFont( calibri, 15 );
        contentStream.setNonStrokingColor(Color.BLACK);
        
        nameWidth = calibri.getStringWidth(c.razon_social) /1000 * 15;
        //System.out.println("nameWidth: " + nameWidth);
        xPosition = (pageWidth - nameWidth)/2;
         
        contentStream.setTextMatrix(new Matrix(1,0,0,1,xPosition,597 ));           
        contentStream.showText(c.razon_social);
        
        contentStream.setFont( calibri, 11 );
        contentStream.setNonStrokingColor(Color.BLACK);
         
        contentStream.setTextMatrix(new Matrix(1,0,0,1,275, (float) 581.4));           
        contentStream.showText(p1.determinante + " " + d.unidad);
        
        contentStream.setTextMatrix(new Matrix(1,0,0,1,235,(float)524.5));           
        contentStream.showText(c.horas_texto);
        
        contentStream.setTextMatrix(new Matrix(1,0,0,1,360,(float)524.5));           
        contentStream.showText(c.fecha_texto_diploma);
        
        contentStream.setFont( calibriBold, 10 ); 
        
        float nameWidthStroked = calibri.getStringWidth(c.nombre_curso) /1000 * 10;
        //System.out.println("nameWidth: " + nameWidthStroked);
        float strokePosition = (pageWidth - nameWidthStroked)/2;
        contentStream.setTextMatrix(new Matrix(1,0,0,1,strokePosition, 540 ));           
        contentStream.showText(c.nombre_curso);
        
        
        
        contentStream.endText();
        
        contentStream.setStrokingColor(Color.BLACK);
        contentStream.moveTo(strokePosition,538);
        contentStream.lineTo(strokePosition + nameWidthStroked + 4, 538);
        contentStream.stroke();
        
        // Make sure that the content stream is closed:
        contentStream.close();
        
        // Save the results and ensure that the document is properly closed:
        


    }
    private static void imprimirDiplomaDoble(Participante p1, Participante p2, Curso c, Directorio d, PDDocument document, PDPage page, PDFont calibri, PDFont calibriBold, PDFont pristina) throws IOException{
        
        /*Directorio d = null;
        
        ListIterator <Directorio> it = listaDirectorio.listIterator();
 
        while (it.hasNext()){
              d = it.next();
             if (d.determinante.equalsIgnoreCase(p1.determinante)){
                break;
             }
        }*/

        COSDictionary pageDict = page.getCOSObject();
        COSDictionary newPageDict = new COSDictionary(pageDict);

        PDPage newPage = new PDPage(newPageDict);
        document.addPage(newPage);
        
        // Start a new content stream which will "hold" the to be created content
        PDPageContentStream contentStream = new PDPageContentStream(document, newPage, true, true);
        
        //PDImageXObject ximage = PDImageXObject.createFromFile("src/files/logo.png", document);
        //contentStream.drawImage(ximage, 100, 10);
                
        float pageWidth = newPage.getMediaBox().getWidth();
        float pageHeight = newPage.getMediaBox().getHeight();
        
        //System.out.println("pageWidth: " + pageWidth + "\npageHeight: " + pageHeight);

        // Print Upper Side
        contentStream.beginText();
        contentStream.setFont( pristina, 28 );
        contentStream.setNonStrokingColor(0,112,192);
        
        float nameWidth = pristina.getStringWidth(p1.nombre + " " + p1.apellidos) /1000 * 28;
        
        //System.out.println("nameWidth: " + nameWidth);
        
        float xPosition = (pageWidth - nameWidth)/2 + 15;
        
        contentStream.setTextMatrix(new Matrix(1,0,0,1,xPosition,622 ));           
        contentStream.showText(p1.nombre + " " + p1.apellidos);
        
        contentStream.setFont( calibri, 15 );
        contentStream.setNonStrokingColor(Color.BLACK);
        
        nameWidth = calibri.getStringWidth(c.razon_social) /1000 * 15;
        //System.out.println("nameWidth: " + nameWidth);
        xPosition = (pageWidth - nameWidth)/2;
         
        contentStream.setTextMatrix(new Matrix(1,0,0,1,xPosition,597 ));           
        contentStream.showText(c.razon_social);
        
        contentStream.setFont( calibri, 11 );
        contentStream.setNonStrokingColor(Color.BLACK);
         
        contentStream.setTextMatrix(new Matrix(1,0,0,1,275, (float) 581.4));           
        contentStream.showText(p1.determinante + " " + d.unidad);
        
        contentStream.setTextMatrix(new Matrix(1,0,0,1,235,(float)524.5));           
        contentStream.showText(c.horas_texto);
        
        contentStream.setTextMatrix(new Matrix(1,0,0,1,360,(float)524.5));           
        contentStream.showText(c.fecha_texto_diploma);
        
        contentStream.setFont( calibriBold, 10 ); 
        
        float nameWidthStroked = calibri.getStringWidth(c.nombre_curso) /1000 * 10;
        //System.out.println("nameWidth: " + nameWidthStroked);
        float strokePosition = (pageWidth - nameWidthStroked)/2;
        contentStream.setTextMatrix(new Matrix(1,0,0,1,strokePosition, 540 ));           
        contentStream.showText(c.nombre_curso);
        
        
        
        contentStream.endText();
        
        contentStream.setStrokingColor(Color.BLACK);
        contentStream.moveTo(strokePosition,538);
        contentStream.lineTo(strokePosition + nameWidthStroked + 4, 538);
        contentStream.stroke();
        
        
        // Print DOWN Side
        /*it = listaDirectorio.listIterator();
 
        while (it.hasNext()){
              d = it.next();
             if (d.determinante.equalsIgnoreCase(p2.determinante)){
                break;
             }
        }*/
        
        contentStream.beginText();
        contentStream.setFont( pristina, 28 );
        contentStream.setNonStrokingColor(0,112,192);
        
        nameWidth = pristina.getStringWidth(p2.nombre + " " + p2.apellidos) /1000 * 28;
        
        //System.out.println("nameWidth: " + nameWidth);
        
        xPosition = (pageWidth - nameWidth)/2 + 15;
        
        contentStream.setTextMatrix(new Matrix(1,0,0,1,xPosition,246 ));          
        contentStream.showText(p2.nombre + " " + p2.apellidos);
        
        contentStream.setFont( calibri, 15 );
        contentStream.setNonStrokingColor(Color.BLACK);
        
        nameWidth = calibri.getStringWidth(c.razon_social) /1000 * 15;
        //System.out.println("nameWidth: " + nameWidth);
        xPosition = (pageWidth - nameWidth)/2;
         
        contentStream.setTextMatrix(new Matrix(1,0,0,1,xPosition,220 ));           
        contentStream.showText(c.razon_social);
        
        contentStream.setFont( calibri, 11 );
        contentStream.setNonStrokingColor(Color.BLACK);
         
        contentStream.setTextMatrix(new Matrix(1,0,0,1,275, (float) 204.4));           
        contentStream.showText(p2.determinante + " " + d.unidad);
        
        contentStream.setTextMatrix(new Matrix(1,0,0,1,235,(float)147.5));           
        contentStream.showText(c.horas_texto);
        
        contentStream.setTextMatrix(new Matrix(1,0,0,1,360,(float)147.5));           
        contentStream.showText(c.fecha_texto_diploma);
        
        contentStream.setFont( calibriBold, 10 ); 
        
        nameWidthStroked = calibri.getStringWidth(c.nombre_curso) /1000 * 10;
        //System.out.println("nameWidth: " + nameWidthStroked);
        strokePosition = (pageWidth - nameWidthStroked)/2;
        contentStream.setTextMatrix(new Matrix(1,0,0,1,strokePosition, 163 ));           
        contentStream.showText(c.nombre_curso);
        
        
        
        contentStream.endText();
        
        contentStream.setStrokingColor(Color.BLACK);
        contentStream.moveTo(strokePosition,161);
        contentStream.lineTo(strokePosition + nameWidthStroked + 4, 161);
        contentStream.stroke();
        
        
        // Make sure that the content stream is closed:
        contentStream.close();
        
        // Save the results and ensure that the document is properly closed:
        //document.save( "DiplomaSoloTest.pdf");
        //document.close();
    }
    private static void imprimirDC3(ArrayList <Participante> listaParticipantes, Curso c, String chkDC3Firmaa, String chkDC3Logoa, String savePath,  Map<String,String> dosc) throws IOException{
        
        ListIterator <Participante> it = listaParticipantes.listIterator();
        Participante p;
        while(it.hasNext()){
            p = it.next();
            if (p.aprovado){
            imprimirDC3_individual(p,c, chkDC3Firmaa, chkDC3Logoa, savePath, dosc);
            }
        }  
          
    }
    private static void imprimirDC3_individual(Participante p,Curso c, String chkDC3Firmaa, String chkDC3Logoa, String savePath, Map<String,String> dosc) throws IOException {
        
        //JOptionPane.showMessageDialog(null, "entrando a imrpimir individual");
        PDDocument document;
        //BufferedInputStream file;
        InputStream file = null;
        
        try{
             file = cdiscisa.Cdiscisa.class.getClassLoader().getResourceAsStream("files/DC3_base_firma.pdf");
        }catch (Exception ex){
            JOptionPane.showMessageDialog(null, "Error al cargar el una forma DC3 \nfile: " + String.valueOf(file) + "\n" + ex.toString());
        }
        document = PDDocument.load(file);
        /*
        if(chkDC3Firmaa.equalsIgnoreCase("true") && chkDC3Logoa.equalsIgnoreCase("true")){
            document = PDDocument.load(new File("src/files/DC3_base_firma.pdf"));
        } else if (chkDC3Firmaa.equalsIgnoreCase("true")){
            document = PDDocument.load(new File("src/files/DC3_base_firma.pdf"));
        } else if (chkDC3Logoa.equalsIgnoreCase("true")){
            document = PDDocument.load(new File("src/files/DC3_base_firma.pdf"));
        } else{
            document = PDDocument.load(new File("src/files/DC3_base_firma.pdf"));
        }   */
         
        
        
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
        contentStream.showText(c.nombre_empresa);
        
        
        char[] rfc_emprea_array = c.rfc_empresa.toCharArray();
        
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
        
        contentStream.setTextMatrix(new Matrix(1,0,0,1,30,356 ));           
        contentStream.showText("6000 SEGURIDAD");
        
        contentStream.setTextMatrix(new Matrix(1,0,0,1,30,332 ));           
        contentStream.showText(c.capacitador);
        
        
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
                charWidth = helvetica_normal.getStringWidth(p.area_tematica) /1000 * 8;
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

        /*
        // 32, 49, 63, 77, 92, 106.5, 120, 134.5, 149, 163, 177, 191.5, 205, 223, 241, 255, 269, 286  
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
        
        contentStream.addRect(50, 750, 500, 100);
        contentStream.setNonStrokingColor(Color.WHITE);
        contentStream.fill();
        contentStream.drawImage(logo, 430,700,150,75);
        contentStream.drawImage(firma, 93,239,72,29);
        */
        
         // Make sure that the content stream is closed:
        contentStream.close();
        Format formatter = new SimpleDateFormat("ddMMMYYYY");
        String formatedDate = formatter.format(c.fecha_inicio);
    
        // Save the results and ensure that the document is properly closed:
        document.save(savePath + File.separator + "p_DC3_" + p.curp + "_" + p.determinante + "_" + p.nombre.replaceAll(" ","_") + "_" + p.apellidos.replaceAll(" ","_") + "_" + formatedDate + ".pdf");
        document.close();
        
        dosc.put(savePath + File.separator + "p_DC3_" + p.curp + "_" + p.determinante + "_" + p.nombre.replaceAll(" ","_") + "_" + p.apellidos.replaceAll(" ","_") + "_" + formatedDate + ".pdf", p.determinante);
        
        
    }
    private static void imprimirConstancias(ArrayList <Participante> listaParticipantes, Curso c, ArrayList <Directorio> listaDirectorio, String chkConstFirma, String chkConstLogo, String savePath, Map<String,String> dosc) throws IOException {
        
        
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
            imprimirUnaConstancia(listaParticipantes, c, d, chkConstFirma, chkConstLogo, savePath, dosc);
        }
      
    }
    private static void imprimirUnaConstancia(ArrayList <Participante> listaParticipantes, Curso c, Directorio d, String chkConstFirma, String chkConstLogo, String savePath, Map<String,String> dosc) throws IOException {
        
        String contanciaTemplate = "";
        
        switch (c.nombre_curso){
            case "PREVENCIÓN Y COMBATE DE INCENDIOS I" : contanciaTemplate = "files/certificado_vacio_incendio_basico.pdf";            
            break;
            case "BUSQUEDA Y RESCATE" : contanciaTemplate = "files/certificado_vacio_busq_rescate.pdf";            
            break;
            case "EVACUACIÓN, BUSQUEDA Y RESCATE" : contanciaTemplate = "files/certificado_vacio_evac_busq_resc.pdf";            
            break;
            case "EVACUACIÓN" : contanciaTemplate = "files/certificado_vacio_evacuacion.pdf";            
            break;
            case "PREVENCIÓN Y COMBATE DE INCENDIOS II" : contanciaTemplate = "files/certificado_vacio_incendio_intermedio.pdf";            
            break;
            case "PREVENCIÓN Y COMBATE DE INCENDIOS III" : contanciaTemplate = "files/certificado_vacio_incendio_avanzado.pdf";            
            break;
            case "FORMACIÓN DE BRIGADAS MULTIFUNCIONALES DE EMERGENCIA" : contanciaTemplate = "files/certificado_vacio_multi.pdf";            
            break;
            case "FORMACIÓN DE BRIGADA MULTIFUNCIONAL DE EMERGENCIA" : contanciaTemplate = "files/certificado_vacio_multi.pdf";            
            break;
            case "PRIMEROS AUXILIOS" : contanciaTemplate = "files/certificado_vacio_primeros_auxilios.pdf";            
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
                
        contentStream.setFont( calibri, 11 );
        contentStream.setNonStrokingColor(Color.BLACK);
        
        contentStream.setTextMatrix(new Matrix(1,0,0,1, 418, 661)); 
        contentStream.showText(df.format(new Date()));
                
                
        contentStream.setFont( calibriBold, 11 );
        
        
        contentStream.setTextMatrix(new Matrix(1,0,0,1, 152 ,(float)590.5)); 
        contentStream.showText(c.razon_social);
        
        contentStream.setTextMatrix(new Matrix(1,0,0,1,140, 577));           
        contentStream.showText(d.determinante + " " + d.unidad);
        
        float charWidth = calibriBold.getStringWidth(d.direccion) /1000 * 11;
        //System.out.println(charWidth + " " + d.direccion.length() + " " +  d.direccion);

        if (charWidth <= 400){
            contentStream.setTextMatrix(new Matrix(1,0,0,1,150, 537));           
            contentStream.showText(d.direccion);
        } else {
            contentStream.setTextMatrix(new Matrix(1,0,0,1,150, 554));           
            contentStream.showText(d.direccion.substring(0, d.direccion.indexOf(" ", 80)));
            contentStream.setTextMatrix(new Matrix(1,0,0,1,150, 539));           
            contentStream.showText(d.direccion.substring(d.direccion.indexOf(" ", 80) + 1, d.direccion.length()));
        }
        
        charWidth = calibriBold.getStringWidth(c.nombre_curso) /1000 * 11;
        
        float pageWidth = page1.getMediaBox().getWidth();
        float xPosition = (pageWidth - charWidth)/2;
                
        contentStream.setTextMatrix(new Matrix(1,0,0,1,xPosition, 500));           
        contentStream.showText(c.nombre_curso);
        
        contentStream.setTextMatrix(new Matrix(1,0,0,1,171, (float)475.5));           
        contentStream.showText(c.fecha_certificado);
        
        contentStream.setTextMatrix(new Matrix(1,0,0,1,171, (float)461.5));           
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
                    contentStream.setTextMatrix(new Matrix(1,0,0,1,180,(float)394.8 - y));           
                    contentStream.showText(p.nombre + " " + p.apellidos);
                    
                    contentStream.setFont( calibri, 11 );
                    
                } else{
                    contentStream.setTextMatrix(new Matrix(1,0,0,1,180, (float)394.8 - y));           
                    contentStream.showText(p.nombre + " " + p.apellidos);
                }
                
                charWidth = calibri.getStringWidth(p.area_puesto) /1000 * 11;
                //System.out.println(charWidth + " " + p.area_puesto);
                
                if (charWidth > 112){
                    contentStream.setFont( calibri, 9 );
                    contentStream.setTextMatrix(new Matrix(1,0,0,1,360, (float)394.8 - y));           
                    contentStream.showText(p.area_puesto);
                
                } else {
                    contentStream.setTextMatrix(new Matrix(1,0,0,1,360, (float)394.8 - y));           
                    contentStream.showText(p.area_puesto);
                }
                
                y = y + (float)13.2;
                 
            }
        }
        
        contentStream.endText();
        
        // Make sure that the content stream is closed:
        contentStream.close();
        
        
        contentStream2.beginText();
        
        contentStream2.setFont( calibri, 11 );
        contentStream2.setNonStrokingColor(Color.BLACK);
        
        contentStream2.setTextMatrix(new Matrix(1,0,0,1, 418, 661)); 
        contentStream2.showText(df.format(new Date()));
                
                
        contentStream2.setFont( calibriBold, 11 );
        
        contentStream2.setTextMatrix(new Matrix(1,0,0,1, 154 ,(float)581.5)); 
        contentStream2.showText(c.razon_social);
        
        contentStream2.setTextMatrix(new Matrix(1,0,0,1,140, 568));           
        contentStream2.showText(d.determinante + " " + d.unidad);
        
        charWidth = calibriBold.getStringWidth(d.direccion) /1000 * 11;
        //System.out.println(charWidth + " " + d.direccion.length() + " " +  d.direccion);

        if (charWidth <= 400){
            contentStream2.setTextMatrix(new Matrix(1,0,0,1,150, 537));           
            contentStream2.showText(d.direccion);
        } else {
            contentStream2.setTextMatrix(new Matrix(1,0,0,1,150, 552));           
            contentStream2.showText(d.direccion.substring(0, d.direccion.indexOf(" ", 80)));
            contentStream2.setTextMatrix(new Matrix(1,0,0,1,150, 539));           
            contentStream2.showText(d.direccion.substring(d.direccion.indexOf(" ", 80) + 1, d.direccion.length()));
        }
        
        charWidth = calibriBold.getStringWidth(c.nombre_curso) /1000 * 11;
        
        pageWidth = page2.getMediaBox().getWidth();
        xPosition = (pageWidth - charWidth)/2;
                
        contentStream2.setTextMatrix(new Matrix(1,0,0,1,xPosition, 487));           
        contentStream2.showText(c.nombre_curso);
        
        contentStream2.setTextMatrix(new Matrix(1,0,0,1,171, 467));           
        contentStream2.showText(c.fecha_certificado);
        
        contentStream2.setTextMatrix(new Matrix(1,0,0,1,171, 454));           
        contentStream2.showText(c.horas_texto);
               
        contentStream2.endText();
        contentStream2.close();
        
        //"Capacitacion_BAE_Centro de Huinala_2631_MULTI_19ago2015"

        //Capacitación + formato tienda + nombre sucursal + numero sucursal + nombre curso + ddmmaaaa
       
        Format formatter = new SimpleDateFormat("ddMMMYYYY");
        String formatedDate = formatter.format(c.fecha_inicio);
    
        // Save the results and ensure that the document is properly closed:
        document.save(savePath + File.separator + "Certificado_" + d.formato + "_" + d.unidad + "_" + d.determinante + "_MULTI_" + formatedDate + ".pdf");
        document.close();
        dosc.put(savePath + File.separator + "Certificado_" + d.formato + "_" + d.unidad + "_" + d.determinante + "_MULTI_" + formatedDate + ".pdf", d.determinante);
        
        
    }
    private static void imprimirDiplomas_main(ArrayList <Participante> listaParticipantes, Curso c, ArrayList <Directorio> listaDirectorio, String chkDipFirma, String chkDipLogo, String savePath, Map<String,String> dosc) throws IOException{
        
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
            imprimirDiplomas(participantesSucursal, c, d, chkDipFirma, chkDipLogo, savePath, dosc);
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
                if(registroPDF.equalsIgnoreCase("files/registro_jorge_2015.pdf")){
                    merge.addSource(cdiscisa.Cdiscisa.class.getClassLoader().getResourceAsStream(registroPDF));
                } else{
                    merge.addSource(new File(registroPDF));
                }
                merge.addSource(new File(lista_auto_pdf));
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
