package emgestion.bdu;

import java.io.File;
import java.math.BigDecimal;
import java.math.RoundingMode;
import java.text.DecimalFormat;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.List;

import javax.swing.JFileChooser;
import javax.swing.JFrame;
import javax.swing.JLabel;
import javax.swing.JOptionPane;
import javax.swing.JPanel;
import javax.swing.JTabbedPane;
import javax.swing.filechooser.FileNameExtensionFilter;

import org.dhatim.fastexcel.Worksheet;
import org.dhatim.fastexcel.reader.CellType;
import org.dhatim.fastexcel.reader.Row;


public class Utilidades {
	private SimpleDateFormat fecha = new SimpleDateFormat("dd/MM/yyyy");
	private SimpleDateFormat fecha2 = new SimpleDateFormat("yyyyMM");
	
	public Utilidades() {
		
	}
	
	public File elegirFichero(JFrame ventana,String tipo, String extension,String titulo) {
		
			File path= new File("");
			// muestra el cuadro de diálogo de archivos, para que el usuario pueda elegir el archivo a abrir
			//FileFilter filtro1 = new FileNameExtensionFilter("Archivos XML", "xml");
			FileNameExtensionFilter filtro = new FileNameExtensionFilter(tipo, extension);
			JFileChooser selectorArchivos = new JFileChooser(path.getAbsolutePath());
			selectorArchivos.setFileSelectionMode(JFileChooser.FILES_AND_DIRECTORIES);
			selectorArchivos.setFileFilter(filtro);
			selectorArchivos.setDialogTitle(titulo);
		
			selectorArchivos.showDialog(ventana,"Aceptar");

			File f1 = selectorArchivos.getSelectedFile(); // obtiene el archivo seleccionado

			// muestra error si es inválido
			if ((f1 == null) ||  !(f1.getName().contains(".xlsx"))) {                     //(f1.getName().equals(""))) {
				JOptionPane.showMessageDialog(ventana, "Archivo inválido o inexistente", "INFO", JOptionPane.INFORMATION_MESSAGE);
				return null;
			} else {
				return f1;
			}
	}
	
	
	public String fet(String f) {
		switch(f) {
			case "VENCIDO":
				return "FETA";
			//break;
			case "INCORPORADO":
				return "FETA_DESFETA";
			//break;
			case "CANCELADO/DESESTIMADO":
				return "FETA";
			//break;
			case "MARGEN":
				return "FETA";
			//break;
		}
		return null;
	}
	
	public String buscaConcepto(String cuenta,String feta) {
		switch (cuenta.substring(0, 3)) {
		
			case "487":
				if(feta=="FETA") {
					return "4332100";
				} else {
					return "5000300";
				}
				
				
			case "709":
				if(feta=="FETA") {
					return "2101400";
				} else {
					return "2151400";
				}
				
				
			default:
				break;
		}
		
		return null;
	}
	
	public int primeraFila(List<Row> datos) {
		for( int x=0;x<datos.size();x++) {
			Row lin=datos.get(x);
			if(lin.getCell(0)!=null) {
				if(lin.getCell(0).getType()== CellType.FORMULA) {
					return x;
				}
			}
			
		}
		
        //JOptionPane.showMessageDialog(this, "No lo ha encontrado, menuda mierda","Información", JOptionPane.INFORMATION_MESSAGE);
		return -1;
	}
	
	public void cabeceraEMGestion(Worksheet hoja) {
		int i=0;
		hoja.value(i, 0,"EMP");
		hoja.value(i, 1,"OFI");
		hoja.value(i, 2,"ENT1");
		hoja.value(i, 3,"PROD1");
		hoja.value(i, 4,"CONTR1");
		hoja.value(i, 5,"ENT2");
		hoja.value(i, 6,"PROD2");
		hoja.value(i, 7,"DESG");
		hoja.value(i, 8,"ENT3");
		hoja.value(i, 9,"PROD3");
		hoja.value(i, 10,"OPER");
		hoja.value(i, 11,"DIV");
		hoja.value(i, 12,"PRODUCTE");
		hoja.value(i, 13,"CONCEPTE");
		hoja.value(i, 14,"VAROPERATIV");
		hoja.value(i, 15,"CODCANAL");
		hoja.value(i, 16,"FECHA CONTABLE");
		hoja.value(i, 17,"PRODBASICO");
		hoja.value(i, 18,"IMPCANTIDAD");
		hoja.value(i, 19,"COMPTE");
		hoja.value(i, 20,"OBSERVA");
		hoja.value(i, 21,"USUARI");
		hoja.value(i, 22,"IMPORTE_CV");
		hoja.value(i, 23,"EMP_OPERA");
		hoja.value(i, 24,"CEN_OPERA");
		hoja.value(i, 25,"ASIENTO");
		hoja.value(i, 26,"feta/desfeta");
		hoja.value(i, 27,"comision");
	}
	
	public void cabeceraBdu(Worksheet hoja) {
		int i=0;
		hoja.value(i, 0,"SECUENCIA_EM");
		hoja.value(i, 1,"MESCONTABLE");
		hoja.value(i, 2,"APLICACION");
		hoja.value(i, 3,"PRODUCTOBASICO");
		hoja.value(i, 4,"VISION");
		hoja.value(i, 5,"ENTIDADCONTRATO");
		hoja.value(i, 6,"PRODUCTOCTRO");
		hoja.value(i, 7,"CONTRATO");;
		hoja.value(i, 8,"ENTIDADCTODESGL");
		hoja.value(i, 9,"PRODCTRODESGL");
		hoja.value(i, 10,"CONTRATODESGL");
		hoja.value(i, 11,"ENTIDADOPERACION");
		hoja.value(i, 12,"PRODOPERACION");
		hoja.value(i, 13,"OPERACION");
		hoja.value(i, 14,"ENTIDADCONTABLE");
		hoja.value(i, 15,"CUENTACONTABLE");
		hoja.value(i, 16,"CENTROCONTABLE");
		hoja.value(i, 17,"DIVISA");
		hoja.value(i, 18,"CONCEPTOGESTION");
		hoja.value(i, 19,"PRODUCTOGESTION");
		hoja.value(i, 20,"CODIGOOPERACION");
		hoja.value(i, 21,"CODIGOIMPORTEKG");
		hoja.value(i, 22,"USUARIO_EM");
		hoja.value(i, 23,"FECHA_EM_CAPTURA");
		hoja.value(i, 24,"COD_TIPOLOGIA");
		hoja.value(i, 25,"TIPOLOGIA");
		hoja.value(i, 26,"IMPORTE");
		hoja.value(i, 27,"IMPORTE_CTV");
		hoja.value(i, 28,"REFUME");
		hoja.value(i, 29,"REFUME2");
		hoja.value(i, 30,"CODIGOOPERABE");
		hoja.value(i, 31,"INDVENCIDOPDTE");
		hoja.value(i, 32,"INDCAPITALINTER");
		hoja.value(i, 33,"CONCEPTOLIQUID");
		hoja.value(i, 34,"FCHULTIMOPAGOCOM");
		hoja.value(i, 35,"IMPCOMISIONES");
		hoja.value(i, 36,"IMPCOSTEDIRECTO");
		hoja.value(i, 37,"PERIODOLIQCOM");
		hoja.value(i, 38,"TIPOPERIODIFCOM");
		hoja.value(i, 39,"PERSONAINV");
		hoja.value(i, 40,"IMPCOMISIONESDIV");
		hoja.value(i, 41,"IMPCOSTEDIRECDIV");
		hoja.value(i, 42,"APLORIGEN");
		hoja.value(i, 43,"PBORIGEN");
		hoja.value(i, 44,"COS");
		hoja.value(i, 45,"T_INDENTMANUAL");
		hoja.value(i, 46,"T_INDRECLASIFICA");
		hoja.value(i, 47,"T_INDARRASTRE");
		hoja.value(i, 48,"T_FCHORIARRASTRE");
	}
	
	public void cabeceraBDIC(Worksheet hoja) {
		int i=0;
		hoja.value(i, 0,"Entidad Bancaria");
		hoja.value(i, 1,"NIF/CIF");
		hoja.value(i, 2,"Fecha de Firma");
		hoja.value(i, 3,"Clave de Banco");
		hoja.value(i, 4,"PAGADO");
		hoja.value(i, 5,"tipo");
		hoja.value(i, 6,"operación");
		hoja.value(i, 7,"TIPO PRODUCTO");
		hoja.value(i, 8,"agrup sector");
		hoja.value(i, 9,"estado operación");
		hoja.value(i, 10,"TOTA LA COMISION");
		hoja.value(i, 11,"Columna1");
		hoja.value(i, 12,"cta cte");
		hoja.value(i, 13,"Columna2");
		hoja.value(i, 14,"Columna3");
		hoja.value(i, 15,"Columna4");
	}

	public void linea(int i,BigDecimal abono, Worksheet hoja, Date fech, String operacion, String situacion, String cuenta,String tipo,String asiento,String comision) {
		String feta=fet(situacion);
		switch (tipo) {
    	case "PR":
    		hoja.value(i, 0,81);
    		hoja.value(i, 1,901);
    		hoja.value(i, 2,"01");
    		hoja.value(i, 3,tipo);
    		hoja.value(i, 4,operacion);
    		hoja.value(i, 5,"01");
    		hoja.value(i, 6,tipo);
    		hoja.value(i, 7,operacion);
    		hoja.value(i, 8,"99");
    		hoja.value(i, 9,"999");
    		hoja.value(i, 10,"999999999999999");
    		hoja.value(i, 11,"EUR");
    		hoja.value(i, 12);
    		//hoja.value(i, 13,buscaConcepto(cuenta,fet(situacion)));
    		hoja.value(i, 13,buscaConcepto(cuenta,feta));
    		hoja.value(i, 14);
    		hoja.value(i, 15);
    		hoja.value(i, 16,fecha.format(fech));
    		hoja.value(i, 17,tipo);
    		hoja.value(i, 18,abono);
    		hoja.value(i, 19,cuenta);
    		hoja.value(i, 20);
    		hoja.value(i, 21,"P097369");
    		hoja.value(i, 22);
    		hoja.value(i, 23,81);
    		hoja.value(i, 24,901);
    		hoja.value(i, 25,asiento);
    		hoja.value(i, 26,feta);
    		hoja.value(i, 27,comision);
    		break;
    	case "KT":
    		hoja.value(i, 0,81);
    		hoja.value(i, 1,901);
    		hoja.value(i, 2,"01");
    		hoja.value(i, 3,"DV");
    		hoja.value(i, 4);
    		hoja.value(i, 5,"99");
    		hoja.value(i, 6,999); 		
    		hoja.value(i, 7,"999999999999999"); 		
    		hoja.value(i, 8,"01");
    		hoja.value(i, 9,tipo);
    		hoja.value(i, 10,operacion);
    		hoja.value(i, 11,"EUR");
    		hoja.value(i, 12);
    		//hoja.value(i, 13,buscaConcepto(cuenta,fet(situacion)));
    		hoja.value(i, 13,buscaConcepto(cuenta,feta));
    		hoja.value(i, 14);
    		hoja.value(i, 15);
    		hoja.value(i, 16,fecha.format(fech));
    		hoja.value(i, 17,"DV");
    		hoja.value(i, 18,abono);
    		hoja.value(i, 19,cuenta);
    		hoja.value(i, 20);
    		hoja.value(i, 21,"P097369");
    		hoja.value(i, 22);
    		hoja.value(i, 23,81);
    		hoja.value(i, 24,901);
    		hoja.value(i, 25,asiento);
    		hoja.value(i, 26,feta);
    		hoja.value(i, 27,comision);
    		break;
    	case "KF":
    		hoja.value(i, 0,81);   		
    		hoja.value(i, 1,901);   		
    		hoja.value(i, 2,"01");   		
    		hoja.value(i, 3,tipo);    		
    		hoja.value(i, 4,operacion);    		
    		hoja.value(i, 5,"99");   		
    		hoja.value(i, 6,"999");
    		hoja.value(i, 7,"999999999999999");  		
    		hoja.value(i, 8,"99");   		
    		hoja.value(i, 9,"999");    		
    		hoja.value(i, 10,"999999999999999");   		
    		hoja.value(i, 11,"EUR");   		
    		hoja.value(i, 12);   		
    		//hoja.value(i, 13,buscaConcepto(cuenta,fet(situacion))); 
    		hoja.value(i, 13,buscaConcepto(cuenta,feta));
    		hoja.value(i, 14);  		
    		hoja.value(i, 15);   		
    		hoja.value(i, 16,fecha.format(fech));		
    		hoja.value(i, 17,tipo);		
    		hoja.value(i, 18,abono);  		
    		hoja.value(i, 19,cuenta);  		
    		hoja.value(i, 20);   		
    		hoja.value(i, 21,"P097369");   		
    		hoja.value(i, 22);   		
    		hoja.value(i, 23,81);    		
    		hoja.value(i, 24,901);    		
    		hoja.value(i, 25,asiento);
    		hoja.value(i, 26,feta);  		
    		hoja.value(i, 27,comision);
    		break;
    	case "XN":
    		hoja.value(i, 0,81);   		
    		hoja.value(i, 1,901);   		
    		hoja.value(i, 2,"01");   		
    		hoja.value(i, 3,tipo);    		
    		hoja.value(i, 4,operacion);    		
    		hoja.value(i, 5,"99");   		
    		hoja.value(i, 6,"999");
    		hoja.value(i, 7,"999999999999999");  		
    		hoja.value(i, 8,"99");   		
    		hoja.value(i, 9,"999");    		
    		hoja.value(i, 10,"999999999999999");   		
    		hoja.value(i, 11,"EUR");   		
    		hoja.value(i, 12);   		
    		//hoja.value(i, 13,buscaConcepto(cuenta,fet(situacion)));
    		hoja.value(i, 13,buscaConcepto(cuenta,feta));
    		hoja.value(i, 14);  		
    		hoja.value(i, 15);   		
    		hoja.value(i, 16,fecha.format(fech));		
    		hoja.value(i, 17,tipo);		
    		hoja.value(i, 18,abono);  		
    		hoja.value(i, 19,cuenta);  		
    		hoja.value(i, 20);   		
    		hoja.value(i, 21,"P097369");   		
    		hoja.value(i, 22);   		
    		hoja.value(i, 23,81);    		
    		hoja.value(i, 24,901);    		
    		hoja.value(i, 25,asiento);
    		hoja.value(i, 26,feta);  		
    		hoja.value(i, 27,comision);
    		break;
    	case "LE":
    		hoja.value(i, 0,81);   		
    		hoja.value(i, 1,901);   		
    		hoja.value(i, 2,"01");   		
    		hoja.value(i, 3,tipo);    		
    		hoja.value(i, 4,operacion);    		
    		hoja.value(i, 5,"99");   		
    		hoja.value(i, 6,"999");
    		hoja.value(i, 7,"999999999999999");  		
    		hoja.value(i, 8,"99");   		
    		hoja.value(i, 9,"999");    		
    		hoja.value(i, 10,"999999999999999");   		
    		hoja.value(i, 11,"EUR");   		
    		hoja.value(i, 12);   		
    		//hoja.value(i, 13,buscaConcepto(cuenta,fet(situacion)));
    		hoja.value(i, 13,buscaConcepto(cuenta,feta));
    		hoja.value(i, 14);  		
    		hoja.value(i, 15);   		
    		hoja.value(i, 16,fecha.format(fech));		
    		hoja.value(i, 17,tipo);		
    		hoja.value(i, 18,abono);  		
    		hoja.value(i, 19,cuenta);  		
    		hoja.value(i, 20);   		
    		hoja.value(i, 21,"P097369");   		
    		hoja.value(i, 22);   		
    		hoja.value(i, 23,81);    		
    		hoja.value(i, 24,901);    		
    		hoja.value(i, 25,asiento);
    		hoja.value(i, 26,feta);  		
    		hoja.value(i, 27,comision);
    		break;
    	default:
    		break;
    }
		
	}
	
	public void lineabdu(int i,BigDecimal abono, Worksheet hoja, Date fech, String operacion, String situacion, String cuenta,String tipo) {
		switch (tipo) {
			case "PR":
				hoja.value(i, 0);
				hoja.value(i, 1,fecha2.format(fech));
				hoja.value(i, 2,tipo);
				hoja.value(i, 3,tipo);
				hoja.value(i, 4,"U");
				hoja.value(i, 5,"01");
				hoja.value(i, 6,tipo);
				hoja.value(i, 7,operacion);;
				hoja.value(i, 8,"01");
				hoja.value(i, 9,tipo);
				hoja.value(i, 10,operacion);
				hoja.value(i, 11,"99");
				hoja.value(i, 12,"999");
				hoja.value(i, 13,"999999999999999");
				hoja.value(i, 14,"0081");
				hoja.value(i, 15,cuenta);
				hoja.value(i, 16,"091");
				hoja.value(i, 17,"EUR");
				hoja.value(i, 18,buscaConcepto(cuenta,fet(situacion)));
				hoja.value(i, 19);
				hoja.value(i, 20);
				hoja.value(i, 21);
				hoja.value(i, 22);
				hoja.value(i, 23);
				hoja.value(i, 24);
				hoja.value(i, 25);
				hoja.value(i, 26,abono);
				hoja.value(i, 27,abono);
				for(int c=operacion.length();c<16;c++) {
					operacion="0"+operacion;
				}
				hoja.value(i, 28,operacion);
				hoja.value(i, 29);
				hoja.value(i, 30,tipo);
				hoja.value(i, 31);
				hoja.value(i, 32);
				hoja.value(i, 33);
				hoja.value(i, 34);
				hoja.value(i, 35);
				hoja.value(i, 36);
				hoja.value(i, 37);
				hoja.value(i, 38);
				hoja.value(i, 39);
				hoja.value(i, 40);
				hoja.value(i, 41);
				hoja.value(i, 42,tipo);
				hoja.value(i, 43,tipo);
				hoja.value(i, 44);
				hoja.value(i, 45);
				hoja.value(i, 46);
				hoja.value(i, 47);
				hoja.value(i, 48);
				break;
			case "KT":
				hoja.value(i, 0);
				hoja.value(i, 1,fecha2.format(fech));
				hoja.value(i, 2,tipo);
				hoja.value(i, 3,tipo);
				hoja.value(i, 4,"U");
				hoja.value(i, 5,"01");
				hoja.value(i, 6,tipo);
				hoja.value(i, 7,operacion);;
				hoja.value(i, 8,"99");
				hoja.value(i, 9,"999");
				hoja.value(i, 10,"999999999999999");
				hoja.value(i, 11,"99");
				hoja.value(i, 12,"999");
				hoja.value(i, 13,operacion);
				hoja.value(i, 14,"0081");
				hoja.value(i, 15,cuenta);
				hoja.value(i, 16,"091");
				hoja.value(i, 17,"EUR");
				hoja.value(i, 18,buscaConcepto(cuenta,fet(situacion)));
				hoja.value(i, 19);
				hoja.value(i, 20);
				hoja.value(i, 21);
				hoja.value(i, 22);
				hoja.value(i, 23);
				hoja.value(i, 24);
				hoja.value(i, 25);
				hoja.value(i, 26,abono);
				hoja.value(i, 27,abono);
				for(int c=operacion.length();c<16;c++) {
					operacion="0"+operacion;
				}
				hoja.value(i, 28,operacion);
				hoja.value(i, 29);
				hoja.value(i, 30,"CR");
				hoja.value(i, 31);
				hoja.value(i, 32);
				hoja.value(i, 33);
				hoja.value(i, 34);
				hoja.value(i, 35);
				hoja.value(i, 36);
				hoja.value(i, 37);
				hoja.value(i, 38);
				hoja.value(i, 39);
				hoja.value(i, 40);
				hoja.value(i, 41,tipo);
				hoja.value(i, 42,tipo);
				hoja.value(i, 43);
				hoja.value(i, 44);
				hoja.value(i, 45);
				hoja.value(i, 46);
				hoja.value(i, 47);
				hoja.value(i, 48);
				break;
			case "KF":
				hoja.value(i, 0);
				hoja.value(i, 1,fecha2.format(fech));
				hoja.value(i, 2,tipo);
				hoja.value(i, 3,tipo);
				hoja.value(i, 4,"U");
				hoja.value(i, 5,"01");
				hoja.value(i, 6,tipo);
				hoja.value(i, 7,operacion);;
				hoja.value(i, 8,"99");
				hoja.value(i, 9,"999");
				hoja.value(i, 10,"999999999999999");
				hoja.value(i, 11,"99");
				hoja.value(i, 12,"999");
				hoja.value(i, 13,"999999999999999");
				hoja.value(i, 14,"0081");
				hoja.value(i, 15,cuenta);
				hoja.value(i, 16,"091");
				hoja.value(i, 17,"EUR");
				hoja.value(i, 18,buscaConcepto(cuenta,fet(situacion)));
				hoja.value(i, 19);
				hoja.value(i, 20);
				hoja.value(i, 21);
				hoja.value(i, 22);
				hoja.value(i, 23);
				hoja.value(i, 24);
				hoja.value(i, 25);
				hoja.value(i, 26,abono);
				hoja.value(i, 27,abono);
				for(int c=operacion.length();c<16;c++) {
					operacion="0"+operacion;
				}
				hoja.value(i, 28,operacion);
				hoja.value(i, 29);
				hoja.value(i, 30,"PS");
				hoja.value(i, 31);
				hoja.value(i, 32);
				hoja.value(i, 33);
				hoja.value(i, 34);
				hoja.value(i, 35);
				hoja.value(i, 36);
				hoja.value(i, 37);
				hoja.value(i, 38);
				hoja.value(i, 39);
				hoja.value(i, 40);
				hoja.value(i, 41);
				hoja.value(i, 42);
				hoja.value(i, 43);
				hoja.value(i, 44);
				hoja.value(i, 45);
				hoja.value(i, 46);
				hoja.value(i, 47,tipo);
				hoja.value(i, 48,tipo);
				break;
			case "XN":
				hoja.value(i, 0);
				hoja.value(i, 1,fecha2.format(fech));
				hoja.value(i, 2,"XF");
				hoja.value(i, 3,"XF");
				hoja.value(i, 4,"U");
				hoja.value(i, 5,"01");
				hoja.value(i, 6,tipo);
				hoja.value(i, 7,operacion);;
				hoja.value(i, 8,"99");
				hoja.value(i, 9,"999");
				hoja.value(i, 10,"999999999999999");
				hoja.value(i, 11,"99");
				hoja.value(i, 12,"999");
				hoja.value(i, 13,"999999999999999");
				hoja.value(i, 14,"0081");
				hoja.value(i, 15,cuenta);
				hoja.value(i, 16,"091");
				hoja.value(i, 17,"EUR");
				hoja.value(i, 18,buscaConcepto(cuenta,fet(situacion)));
				hoja.value(i, 19);
				hoja.value(i, 20);
				hoja.value(i, 21);
				hoja.value(i, 22);
				hoja.value(i, 23);
				hoja.value(i, 24);
				hoja.value(i, 25);
				hoja.value(i, 26,abono);
				hoja.value(i, 27,abono);
				for(int c=operacion.length();c<16;c++) {
					operacion="0"+operacion;
				}
				hoja.value(i, 28,operacion);
				hoja.value(i, 29);
				hoja.value(i, 30,"EP");
				hoja.value(i, 31);
				hoja.value(i, 32);
				hoja.value(i, 33);
				hoja.value(i, 34,"XF");
				hoja.value(i, 35,"XF");
				hoja.value(i, 36);
				hoja.value(i, 37);
				hoja.value(i, 38);
				hoja.value(i, 39);
				hoja.value(i, 40);
				hoja.value(i, 41);
				hoja.value(i, 42);
				hoja.value(i, 43);
				hoja.value(i, 44);
				hoja.value(i, 45);
				hoja.value(i, 46);
				hoja.value(i, 47);
				hoja.value(i, 48);
				break;
			case "LE":
				hoja.value(i, 0);
				hoja.value(i, 1,fecha2.format(fech));
				hoja.value(i, 2,"LE");
				hoja.value(i, 3,"LE");
				hoja.value(i, 4,"U");
				hoja.value(i, 5,"01");
				hoja.value(i, 6,"LE");
				hoja.value(i, 7,operacion);;
				hoja.value(i, 8,"01");
				hoja.value(i, 9,"LE");
				hoja.value(i, 10,operacion);
				hoja.value(i, 11,"99");
				hoja.value(i, 12,"999");
				hoja.value(i, 13,"999999999999999");
				hoja.value(i, 14,"0081");
				hoja.value(i, 15,cuenta);
				hoja.value(i, 16,"0901");
				hoja.value(i, 17,"EUR");
				hoja.value(i, 18,buscaConcepto(cuenta,fet(situacion)));
				hoja.value(i, 19);
				hoja.value(i, 20);
				hoja.value(i, 21);
				hoja.value(i, 22);
				hoja.value(i, 23);
				hoja.value(i, 24);
				hoja.value(i, 25);
				hoja.value(i, 26,abono);
				hoja.value(i, 27,abono);
				//hoja.value(i, 27,abono.setScale(0, RoundingMode.HALF_UP));
				for(int c=operacion.length();c<16;c++) {
					operacion="0"+operacion;
				}
				hoja.value(i, 28,operacion);
				hoja.value(i, 29);
				hoja.value(i, 30,"LS");
				hoja.value(i, 31);
				hoja.value(i, 32);
				hoja.value(i, 33);
				hoja.value(i, 34);
				hoja.value(i, 35);
				hoja.value(i, 36);
				hoja.value(i, 37);
				hoja.value(i, 38);
				hoja.value(i, 39,"LE");
				hoja.value(i, 40,"LE");
				hoja.value(i, 41);
				hoja.value(i, 42);
				hoja.value(i, 43);
				hoja.value(i, 44);
				hoja.value(i, 45);
				hoja.value(i, 46);
				hoja.value(i, 47);
				hoja.value(i, 48);
				break;
		}
	}
	

	
	/**
	 * 
	 * @param lista lista de movimientos
	 * @param anos años que hay
	 * @param modo 0-comisiones     1-periodificaciones
	 */
	public JTabbedPane resu(List<Row> lista, int anos,int modo) {
		JTabbedPane tb=new JTabbedPane();
		if(lista.size()>0) {
			BigDecimal devengo=new BigDecimal(0);
			int textoAno;
			int textoTipo;
			
			Row m;
			for(int a=0;a<anos;a++) {
				JPanel jp=new JPanel();
				BigDecimal resultadoPRV=new BigDecimal(0);
				BigDecimal resultadoKTV=new BigDecimal(0);
				BigDecimal resultadoKFV=new BigDecimal(0);
				BigDecimal resultadoXNV=new BigDecimal(0);
				BigDecimal resultadoLEV=new BigDecimal(0);
				BigDecimal totalV=new BigDecimal(0);
				BigDecimal resultadoPRC=new BigDecimal(0);
				BigDecimal resultadoKTC=new BigDecimal(0);
				BigDecimal resultadoKFC=new BigDecimal(0);
				BigDecimal resultadoXNC=new BigDecimal(0);
				BigDecimal resultadoLEC=new BigDecimal(0);
				BigDecimal totalC=new BigDecimal(0);
				BigDecimal resultadoPRI=new BigDecimal(0);
				BigDecimal resultadoKTI=new BigDecimal(0);
				BigDecimal resultadoKFI=new BigDecimal(0);
				BigDecimal resultadoXNI=new BigDecimal(0);
				BigDecimal resultadoLEI=new BigDecimal(0);
				BigDecimal totalI=new BigDecimal(0);
				for(int i=0;i<lista.size();i++) {
					m=lista.get(i);
					if (modo==0) {
						textoAno=6;
						textoTipo=12;
						devengo=m.getCell(5).asNumber();
					} else {
						textoAno=7;
						textoTipo=21;
						devengo=new BigDecimal(m.getCellText(17)).setScale(2, RoundingMode.HALF_UP);
					}
					if(m.getCellText(textoAno).contains(String.valueOf(a+1))) {
						//BigDecimal devengo=new BigDecimal(m.getCellText(17)).setScale(2, RoundingMode.HALF_UP);
						//Vencidos
						if(m.getCellText(textoTipo).contains("KT-O.S.R-VENCIDO")) {
							resultadoKTV=resultadoKTV.add(m.getCell(5).asNumber());
						}
						if(m.getCellText(textoTipo).contains("PR-O.S.R-VENCIDO")) {
							resultadoPRV=resultadoPRV.add(m.getCell(5).asNumber());
						}
						if(m.getCellText(textoTipo).contains("KF-O.S.R-VENCIDO")) {
							resultadoKFV=resultadoKFV.add(m.getCell(5).asNumber());
						}
						if(m.getCellText(textoTipo).contains("XN-O.S.R-VENCIDO")) {
							resultadoXNV=resultadoXNV.add(m.getCell(5).asNumber());
						}
						if(m.getCellText(textoTipo).contains("LE-O.S.R-VENCIDO")) {
							resultadoLEV=resultadoLEV.add(m.getCell(5).asNumber());
						}
						//Cancelados
						if(m.getCellText(textoTipo).contains("KT-O.S.R-CANCELADO/DESESTIMADO")) {
							resultadoKTC=resultadoKTC.add(m.getCell(5).asNumber());
						}
						if(m.getCellText(textoTipo).contains("PR-O.S.R-CANCELADO/DESESTIMADO")) {
							resultadoPRC=resultadoPRC.add(m.getCell(5).asNumber());
						}
						if(m.getCellText(textoTipo).contains("KF-O.S.R-CANCELADO/DESESTIMADO")) {
							resultadoKFC=resultadoKFC.add(m.getCell(5).asNumber());
						}
						if(m.getCellText(textoTipo).contains("XN-O.S.R-CANCELADO/DESESTIMADO")) {
							resultadoXNC=resultadoXNC.add(m.getCell(5).asNumber());
						}
						if(m.getCellText(textoTipo).contains("LE-O.S.R-CANCELADO/DESESTIMADO")) {
							resultadoLEC=resultadoLEC.add(m.getCell(5).asNumber());
						}
						if(m.getCellText(textoTipo).contains("KT-O.S.R-MARGEN")) {
							resultadoKTC=resultadoKTC.add(m.getCell(5).asNumber());
						}
						if(m.getCellText(textoTipo).contains("PR-O.S.R-MARGEN")) {
							resultadoPRC=resultadoPRC.add(m.getCell(5).asNumber());
						}
						if(m.getCellText(textoTipo).contains("KF-O.S.R-MARGEN")) {
							resultadoKFC=resultadoKFC.add(m.getCell(5).asNumber());
						}
						if(m.getCellText(textoTipo).contains("XN-O.S.R-MARGEN")) {
							resultadoXNC=resultadoXNC.add(m.getCell(5).asNumber());
						}
						if(m.getCellText(textoTipo).contains("LE-O.S.R-MARGEN")) {
							resultadoLEC=resultadoLEC.add(m.getCell(5).asNumber());
						}
						//Incorporados
						if(m.getCellText(textoTipo).contains("KT-O.S.R-INCORPORADO")) {
							resultadoKTI=resultadoKTI.add(devengo);
						}
						if(m.getCellText(textoTipo).contains("PR-O.S.R-INCORPORADO")) {
							resultadoPRI=resultadoPRI.add(devengo);
						}
						if(m.getCellText(textoTipo).contains("KF-O.S.R-INCORPORADO")) {
							resultadoKFI=resultadoKFI.add(devengo);
						}
						if(m.getCellText(textoTipo).contains("XN-O.S.R-INCORPORADO")) {
							resultadoXNI=resultadoXNI.add(devengo);
						}
						if(m.getCellText(textoTipo).contains("LE-O.S.R-INCORPORADO")) {
							resultadoLEI=resultadoLEI.add(devengo);
						}
					}
				}
				totalV=totalV.add(resultadoXNV).add(resultadoKFV).add(resultadoKTV).add(resultadoPRV).add(resultadoLEV);
				totalC=totalC.add(resultadoXNC).add(resultadoKFC).add(resultadoKTC).add(resultadoPRC).add(resultadoLEC);
				totalI=totalI.add(resultadoXNI).add(resultadoKFI).add(resultadoKTI).add(resultadoPRI).add(resultadoLEI);
				
				DecimalFormat formato1 = new DecimalFormat("#.00");
				JLabel jl=new JLabel("<html><body>VENCIDO<br>Resultado KT: "+formato1.format(resultadoKTV)+"<br>ResultadoPR: "+formato1.format(resultadoPRV)+"<br>ResultadoKF: "+formato1.format(resultadoKFV)+"<br>ResultadoXN: "+formato1.format(resultadoXNV)+"<br>ResultadoLE: "+formato1.format(resultadoLEV)+"<br>Total: "+formato1.format(totalV)
				+"<br>CANCELADO<br>Resultado KT: "+formato1.format(resultadoKTC)+"<br>ResultadoPR: "+formato1.format(resultadoPRC)+"<br>ResultadoKF: "+formato1.format(resultadoKFC)+"<br>ResultadoXN: "+formato1.format(resultadoXNC)+"<br>ResultadoLE: "+formato1.format(resultadoLEC)+"<br>Total: "+formato1.format(totalC)
				+"<br>INCORPORADO<br>Resultado KT: "+formato1.format(resultadoKTI)+"<br>ResultadoPR: "+formato1.format(resultadoPRI)+"<br>ResultadoKF: "+formato1.format(resultadoKFI)+"<br>ResultadoXN: "+formato1.format(resultadoXNI)+"<br>ResultadoLE: "+formato1.format(resultadoLEI)+"<br>Total: "+formato1.format(totalI)+"</body></html>");
				jp.add(jl);
				tb.addTab("Resultados Año "+(a+1), jp);		}
		}/*else {
			JOptionPane.showMessageDialog(null, "No hay movimientos periodificación,tonto el culo","Información", JOptionPane.INFORMATION_MESSAGE);
		}*/
		return tb;
	}
}
