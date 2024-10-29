package emgestion.bdu;

import java.awt.BorderLayout;
import java.awt.GridLayout;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.awt.event.WindowAdapter;
import java.awt.event.WindowEvent;
import java.io.File;
import java.io.FileOutputStream;
import java.math.BigDecimal;
import java.math.RoundingMode;
import java.util.ArrayList;
import java.util.Calendar;
import java.util.Date;
import java.util.List;
import java.util.concurrent.ExecutionException;
import javax.swing.JButton;
import javax.swing.JDialog;
import javax.swing.JFrame;
import javax.swing.JMenu;
import javax.swing.JMenuBar;
import javax.swing.JMenuItem;
import javax.swing.JOptionPane;
import javax.swing.JPanel;
import javax.swing.JProgressBar;
import javax.swing.JScrollPane;
import javax.swing.JTabbedPane;
import org.dhatim.fastexcel.Workbook;
import org.dhatim.fastexcel.Worksheet;
import org.dhatim.fastexcel.reader.Row;
import com.toedter.calendar.JDateChooser;

public class Emgbdu extends JFrame{
	
	private static final long serialVersionUID = 9173863219222483394L;
	//private JTextArea area = new JTextArea();
	private JPanel area=new JPanel(new BorderLayout());
	private Date fech,fech2;
	private Utilidades uti=new Utilidades();
	private List<Row> comisiones= new ArrayList<Row>();
	private List<Row> periodificaciones= new ArrayList<Row>();
	private List<Integer> anos=new ArrayList<Integer>();

	public Emgbdu() {		
		super("Contabilidad Paula");
		this.setExtendedState(MAXIMIZED_BOTH);
		//this.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
		this.setDefaultCloseOperation(JFrame.DO_NOTHING_ON_CLOSE);
		this.addWindowListener(new WindowAdapter() {
			@Override
			public void windowClosing(WindowEvent e) {				          
					salida();
			}
		});
		JScrollPane scroll =new JScrollPane(area);
		
		this.getContentPane().add(scroll);

		anos.add(0);
		anos.add(0);
		
		JMenuBar menu= new JMenuBar();
		JMenu inicio= new JMenu("Inicio");
		JMenuItem cargar = new JMenuItem("Nuevo");
		cargar.addActionListener(new ActionListener(){
	         public void actionPerformed (ActionEvent e)
	         {
	        	 nuevo();
	         }
		});
		
		JMenuItem salir = new JMenuItem("Salir");
		salir.addActionListener(new ActionListener(){
	         public void actionPerformed (ActionEvent e)
	         {
	            salida();
	         }
	       });
		
		inicio.add(cargar);
		inicio.add(salir);
		menu.add(inicio);
		
		setJMenuBar(menu);		
		
		this.setVisible(true);

		proceso();
	}
	
	private void nuevo() {
		comisiones.clear();
		periodificaciones.clear();
		area.removeAll();
		area.updateUI();
		anos.set(0, 0);
		anos.set(1, 0);
		proceso();
	}

	private void proceso() {
		CargadorFicheros tc=null;
		CargadorFicheros tp=null;
		fechas();
		File f1=uti.elegirFichero(this,"Excel","xlsx","Elegir fichero comisiones---Cancelar si no hay");
 		if (f1!=null) {			
 			try {
				tc=pru(f1,comisiones,"Comisiones");
			} catch (InterruptedException | ExecutionException e1) {
				JOptionPane.showMessageDialog(null, e1.toString(),"Error salida", JOptionPane.ERROR_MESSAGE);
			}
 		}
 		if(tc!=null) {
 			while(!tc.isDone()) {
 			}
 		}

 		f1=uti.elegirFichero(this,"Excel","xlsx","Elegir fichero Periodificaciones---Cancelar si no hay");
 		if (f1!=null) {
 			try {
 				tp=pru(f1,periodificaciones,"comis");
			} catch (InterruptedException | ExecutionException e1) {
				JOptionPane.showMessageDialog(null, e1.toString(),"Error salida", JOptionPane.ERROR_MESSAGE);
			}
 		}
 		if(tp!=null) {
 			while(!tp.isDone()) {
 			}
 		}
 
 		
 		
 		area.add(resuls(),BorderLayout.CENTER);
 		JButton bot=new JButton("Crear fichero");
 		bot.addActionListener(new ActionListener(){
	         public void actionPerformed (ActionEvent e)
	         {
	        	 guardar();
	        	 guardarBDIC();
	        	 //guardarPrueba();
	         }
	       });
 		area.add(bot,BorderLayout.SOUTH);
 		area.updateUI();
	}
	
	private void depurar(Worksheet hojaADepurar, Worksheet depurada) {
		
		List<String> visitados= new ArrayList<String>();
		List<String> repetidos= new ArrayList<String>();
		
		
		int f=1;
		int orden=1;
		
		//busco los que no estan repetidos, que guardo en visitados y los repetidos, que guardo en repetidos
		String buscar=(String)hojaADepurar.value(f, 7)+(String)hojaADepurar.value(f, 15)+(String)hojaADepurar.value(f, 18);
		while(hojaADepurar.value(f, 7)!=null) {
			if(!visitados.contains(buscar)) {
				visitados.add(buscar);
			} else {
				if(!repetidos.contains(buscar)) {
					repetidos.add(buscar);
				}
			}
			f++;
			buscar=(String)hojaADepurar.value(f, 7)+(String)hojaADepurar.value(f, 15)+(String)hojaADepurar.value(f, 18);
		}
		for(int i=0;i<repetidos.size();i++) {
			visitados.remove(repetidos.get(i));
		}
	
		//trato los que no estan repetidos y los paso a la bdu depurada
		for(int i=0;i<visitados.size();i++) {
			f=1;
			String buscado=visitados.get(i);
			buscar=(String)hojaADepurar.value(f, 7)+(String)hojaADepurar.value(f, 15)+(String)hojaADepurar.value(f, 18);
			while(hojaADepurar.value(f, 7)!=null) {
				if(buscar.equals(buscado)) {
					for(int x=0;x<44;x++) {
						
						if(x>25&&x<28) {
							depurada.value(orden, x,(BigDecimal)hojaADepurar.value(f,x));
						} else {
							depurada.value(orden, x,(String)hojaADepurar.value(f,x));
						}
					}
					orden++;
					break;
				}
				f++;
				buscar=(String)hojaADepurar.value(f, 7)+(String)hojaADepurar.value(f, 15)+(String)hojaADepurar.value(f, 18);
			}
		}
		
		//trato los repetidos y los paso a la dbu depurada
		for(int i=0;i<repetidos.size();i++) {
			f=1;
			int fe=0;
			String bus=repetidos.get(i);
			boolean encontrado=false;
			BigDecimal cantidad=new BigDecimal(0);
			buscar=(String)hojaADepurar.value(f, 7)+(String)hojaADepurar.value(f, 15)+(String)hojaADepurar.value(f, 18);
			while(hojaADepurar.value(f, 7)!=null) {
				//codigo por hacer
				if(buscar.equals(bus)) {
					if(!encontrado) {
						cantidad=(BigDecimal)hojaADepurar.value(f, 26);
						//System.out.println("0 "+cantidad);
						fe=f;
						encontrado=true;
					} else {
						//System.out.println("1 "+cantidad);
						cantidad=cantidad.add((BigDecimal)hojaADepurar.value(f, 26));
						//System.out.println("1.5 "+(BigDecimal)hojaADepurar.value(f, 26));
						//System.out.println("2 "+cantidad);
					}
				}
				f++;
				buscar=(String)hojaADepurar.value(f, 7)+(String)hojaADepurar.value(f, 15)+(String)hojaADepurar.value(f, 18);
			}
			if(cantidad.compareTo(new BigDecimal(0))!=0) {
				for(int x=0;x<44;x++) {
				
					if(x>25&&x<28) {
						depurada.value(orden, x,cantidad);
					} else {
						depurada.value(orden, x,(String)hojaADepurar.value(fe,x));
					}
				}
			} else {
				orden--;
			}
			orden++;
			
		}
		
		
	}

	private JTabbedPane resuls() {
		JTabbedPane princi=new JTabbedPane();
		princi.addTab("Comisiones", uti.resu(comisiones, anos.get(0), 0));
		princi.addTab("Periodificaciones", uti.resu(periodificaciones, anos.get(1), 1));
		return princi;
	}

	
	private CargadorFicheros pru(File f,List<Row> lista,String texto) throws InterruptedException, ExecutionException {
		JDialog ven=new JDialog(this,"Cargando fichero, espere...");
		JProgressBar jb=new JProgressBar(0,100);
		ven.setBounds(300, 300, 300, 100);
		JPanel fondo=new JPanel(new GridLayout(3,1));
		//ven.setLayout(new GridLayout());
		jb.setBounds(40,40,300,100);         
		jb.setValue(0);    
		jb.setStringPainted(true);
		fondo.add(jb);
		ven.add(fondo);
		//ven.pack();
		ven.setVisible(true);
		CargadorFicheros trab=new CargadorFicheros(this,ven,fondo,jb,f,lista,texto,anos);
		trab.execute();
		return trab;
	}
	

	
	
	private void guardarBDIC() {
		try {
			//final String nombreArchivo = "ComisionesRepercutidas.xlsx";
			String nombreArchivo =JOptionPane.showInputDialog("Nombre del archivo de BDIC: ");
			if(nombreArchivo==null) {
				nombreArchivo="BDIC.xlsx";
			} else {
				nombreArchivo+=".xlsx";
			}
			File directorioActual = new File(".");
			String ubicacion = directorioActual.getAbsolutePath();
			String ubicacionArchivoSalida = ubicacion.substring(0, ubicacion.length() - 1) + nombreArchivo;
			FileOutputStream outputStream;
			outputStream = new FileOutputStream(ubicacionArchivoSalida);
			@SuppressWarnings("resource")
			Workbook salida=new Workbook(outputStream,this.getTitle(),null);
			String MES[] = {"ENERO", "FEBRERO", "MARZO", "ABRIL", "MAYO", "JUNIO", "JULIO", "AGOSTO", "SEPTIEMBRE", "OCTUBRE", "NOVIEMBRE", "DICIEMBRE"};
			Calendar calendario=Calendar.getInstance();
			calendario.setTime(fech);
			String mesActual=MES[calendario.get(Calendar.MONTH)];
			Worksheet hoja=salida.newWorksheet("BDIC period "+mesActual.toUpperCase());
			Worksheet hoja1=salida.newWorksheet("BDIC real "+mesActual.toUpperCase());
		
			uti.cabeceraBDIC(hoja);
			uti.cabeceraBDIC(hoja1);
			
			Row lin;
			int l=1;
			for(int i=0;i<comisiones.size();i++) {		
				lin=comisiones.get(i);	
				if(lin.getCellText(13).contains("709")) {
					String fra="COMISION VENCIDA AÑO "+lin.getCellText(6).substring(lin.getCellText(6).length()-1);
	
					hoja1.value(l, 0,"");
					hoja1.value(l, 1,"");
					hoja1.value(l, 2,"");
					hoja1.value(l, 3,"");
					hoja1.value(l, 4,"");
					hoja1.value(l, 5,"");
					hoja1.value(l, 6,lin.getCellText(7));
					hoja1.value(l, 7,lin.getCellText(8));
					hoja1.value(l, 8,lin.getCellText(10));
					hoja1.value(l, 9, lin.getCellText(11));
					hoja1.value(l, 10,lin.getCell(5).asNumber().setScale(2, RoundingMode.HALF_UP));
					hoja1.value(l, 11,lin.getCellText(15));
					hoja1.value(l, 12,lin.getCellText(13));
					hoja1.value(l, 13,"");
					hoja1.value(l, 14,"");
					hoja1.value(l, 15,fra);
					l++;
					}
			}
			int lp=1;
			for(int i=0;i<periodificaciones.size();i++) {
				
				lin=periodificaciones.get(i);
				
				String cuenta=lin.getCellText(22);
				if(cuenta.equalsIgnoreCase("0")) {
					cuenta=lin.getCellText(28);
				}
				
				if(cuenta.contains("709")) {
					BigDecimal devengo=new BigDecimal(lin.getCellText(17)).setScale(2, RoundingMode.HALF_UP);
					BigDecimal devengo2=new BigDecimal(lin.getCellText(5)).setScale(2, RoundingMode.HALF_UP);
					String frase=lin.getCellText(24);
					if(frase.equalsIgnoreCase("0")) {
						frase=lin.getCellText(30);
					}
					String fra2="PERIODIFICACION AÑO "+lin.getCellText(7).substring(lin.getCellText(7).length()-1);
					
					if(lin.getCellText(20).contains("INCORPORADO")) {
						hoja.value(lp, 0,"");
						hoja.value(lp, 1,"");
						hoja.value(lp, 2,"");
						hoja.value(lp, 3,"");
						hoja.value(lp, 4,"");
						hoja.value(lp, 5,"");
						hoja.value(lp, 6,lin.getCellText(8));
						hoja.value(lp, 7,lin.getCellText(9));
						hoja.value(lp, 8,lin.getCellText(10));
						hoja.value(lp, 9, lin.getCellText(20));
						hoja.value(lp, 10,devengo.setScale(2, RoundingMode.HALF_UP));
						hoja.value(lp, 11,frase);
						hoja.value(lp, 12,cuenta);
						hoja.value(lp, 13,"");
						hoja.value(lp, 14,"");
						hoja.value(lp, 15,fra2);
						lp++;
					} else {
						if(lin.getCellText(20).contains("VENCID")) {
							fra2="PERIODI COMISION VENCIDA AÑO "+lin.getCellText(7).substring(lin.getCellText(7).length()-1);
						} else {
							fra2="PERIODI COMISION CANCELADA AÑO "+lin.getCellText(7).substring(lin.getCellText(7).length()-1);
						}
						hoja1.value(l, 0,"");
						hoja1.value(l, 1,"");
						hoja1.value(l, 2,"");
						hoja1.value(l, 3,"");
						hoja1.value(l, 4,"");
						hoja1.value(l, 5,"");
						hoja1.value(l, 6,lin.getCellText(8));
						hoja1.value(l, 7,lin.getCellText(9));
						hoja1.value(l, 8,lin.getCellText(10));
						hoja1.value(l, 9, lin.getCellText(20));
						hoja1.value(l, 10,devengo2.setScale(2, RoundingMode.HALF_UP));
						hoja1.value(l, 11,frase);
						hoja1.value(l, 12,cuenta);
						hoja1.value(l, 13,"");
						hoja1.value(l, 14,"");
						hoja1.value(l, 15,fra2);
						l++;
					}
				}
				
			}
			
			
			salida.finish();			
			outputStream.close();
			
            
            JOptionPane.showMessageDialog(this, "Fichero creado","Información", JOptionPane.INFORMATION_MESSAGE);
		} catch (Exception e) {
			JOptionPane.showMessageDialog(this, e.toString(),"Error salida", JOptionPane.ERROR_MESSAGE);	
		}
	}
	
	private void guardar() {
		try {
			//final String nombreArchivo = "ComisionesRepercutidas.xlsx";
			String nombreArchivo =JOptionPane.showInputDialog("Nombre del archivo de BDU: ");
			if(nombreArchivo==null) {
				nombreArchivo="ComisionesRepercutidas.xlsx";
			} else {
				nombreArchivo+=".xlsx";
			}
			File directorioActual = new File(".");
			String ubicacion = directorioActual.getAbsolutePath();
			String ubicacionArchivoSalida = ubicacion.substring(0, ubicacion.length() - 1) + nombreArchivo;
			FileOutputStream outputStream;
			outputStream = new FileOutputStream(ubicacionArchivoSalida);
			@SuppressWarnings("resource")
			Workbook salida=new Workbook(outputStream,this.getTitle(),null);
			Worksheet hoja=salida.newWorksheet("EMG");
			Worksheet hoja1=salida.newWorksheet("KT");
			Worksheet hoja2=salida.newWorksheet("PR");
			Worksheet hoja3=salida.newWorksheet("XN");
			Worksheet hoja4=salida.newWorksheet("KF");
			Worksheet hoja5=salida.newWorksheet("LE");
			Worksheet hoja6=salida.newWorksheet("KTU");
			Worksheet hoja7=salida.newWorksheet("PRU");
			Worksheet hoja8=salida.newWorksheet("XNU");
			Worksheet hoja9=salida.newWorksheet("KFU");
			Worksheet hoja10=salida.newWorksheet("LEU");
			uti.cabeceraEMGestion(hoja);
			uti.cabeceraBdu(hoja1);
			uti.cabeceraBdu(hoja2);
			uti.cabeceraBdu(hoja3);
			uti.cabeceraBdu(hoja4);
			uti.cabeceraBdu(hoja5);
			uti.cabeceraBdu(hoja6);
			uti.cabeceraBdu(hoja7);
			uti.cabeceraBdu(hoja8);
			uti.cabeceraBdu(hoja9);
			uti.cabeceraBdu(hoja10);
			Row lin;
			int l=1;
			int lPR=1;
			int lKF=1;
			int lKT=1;
			int lXN=1;
			int lLE=1;
			String asiento,comision;
			for(int i=0;i<comisiones.size();i++) {
				lin=comisiones.get(i);
				asiento="Comisión FETA";
				comision="Comisión "+lin.getCellText(11)+ " año ";
				//System.out.println(comision);
				uti.linea(l,lin.getCell(5).asNumber().setScale(2, RoundingMode.HALF_UP), hoja,fech,lin.getCellText(7),"MARGEN",lin.getCellText(13),lin.getCellText(8),asiento,comision+lin.getCellText(6).substring(lin.getCellText(6).length() - 1));
				l++;
				if(lin.getCellText(13).contains("487")) {
					switch(lin.getCellText(8)) {
						case "KT":
							uti.lineabdu(lKT,lin.getCell(5).asNumber().setScale(2, RoundingMode.HALF_UP), hoja1,fech,lin.getCellText(7),"MARGEN",lin.getCellText(13),lin.getCellText(8));
							lKT++;
							break;
						case "PR":
							uti.lineabdu(lPR,lin.getCell(5).asNumber().setScale(2, RoundingMode.HALF_UP), hoja2,fech,lin.getCellText(7),"MARGEN",lin.getCellText(13),lin.getCellText(8));
							lPR++;
							break;
						case "XN":
							uti.lineabdu(lXN,lin.getCell(5).asNumber().setScale(2, RoundingMode.HALF_UP), hoja3,fech,lin.getCellText(7),"MARGEN",lin.getCellText(13),lin.getCellText(8));
							lXN++;
							break;
						case "KF":
							uti.lineabdu(lKF,lin.getCell(5).asNumber().setScale(2, RoundingMode.HALF_UP), hoja4,fech,lin.getCellText(7),"MARGEN",lin.getCellText(13),lin.getCellText(8));
							lKF++;
							break;
						case "LE":
							uti.lineabdu(lLE,lin.getCell(5).asNumber().setScale(2, RoundingMode.HALF_UP), hoja5,fech,lin.getCellText(7),"MARGEN",lin.getCellText(13),lin.getCellText(8));
							lLE++;
							break;
					}
				}
			}
			
			for(int i=0;i<periodificaciones.size();i++) {
				
				lin=periodificaciones.get(i);
				
				//    abono-5  devengado-17 operacion-8 situacion-20 cuentaIngreso-22  cuentaResta-25  tipo-9
				switch(uti.fet(lin.getCellText(20))) {
				case "FETA":
					asiento="Periodificación FETA/Asiento en firme";
					comision="Periodificación "+lin.getCellText(20)+ " año ";
					String cuent1=lin.getCellText(22);
					if(cuent1.equals("0")) {
						//System.out.println(cuent1);
						cuent1=lin.getCellText(28);
						//System.out.println(cuent1);
					}
					//uti.linea(l,lin.getCell(5).asNumber().setScale(2, RoundingMode.HALF_UP), hoja,fech,lin.getCellText(8),lin.getCellText(20),lin.getCellText(22),lin.getCellText(9),asiento,comision+lin.getCellText(7).substring(lin.getCellText(7).length() - 1));
					uti.linea(l,lin.getCell(5).asNumber().setScale(2, RoundingMode.HALF_UP), hoja,fech,lin.getCellText(8),lin.getCellText(20),cuent1,lin.getCellText(9),asiento,comision+lin.getCellText(7).substring(lin.getCellText(7).length() - 1));
					l++;
					if(lin.getCellText(22).contains("487")) {
						switch(lin.getCellText(9)) {
							case "KT":
								uti.lineabdu(lKT,lin.getCell(5).asNumber().setScale(2, RoundingMode.HALF_UP), hoja1,fech,lin.getCellText(8),lin.getCellText(20),lin.getCellText(22),lin.getCellText(9));
								lKT++;
								break;
							case "PR":
								uti.lineabdu(lPR,lin.getCell(5).asNumber().setScale(2, RoundingMode.HALF_UP), hoja2,fech,lin.getCellText(8),lin.getCellText(20),lin.getCellText(22),lin.getCellText(9));
								lPR++;
								break;
							case "XN":
								uti.lineabdu(lXN,lin.getCell(5).asNumber().setScale(2, RoundingMode.HALF_UP), hoja3,fech,lin.getCellText(8),lin.getCellText(20),lin.getCellText(22),lin.getCellText(9));
								lXN++;
								break;
							case "KF":
								uti.lineabdu(lKF,lin.getCell(5).asNumber().setScale(2, RoundingMode.HALF_UP), hoja4,fech,lin.getCellText(8),lin.getCellText(20),lin.getCellText(22),lin.getCellText(9));
								lKF++;
								break;
							case "LE":
								uti.lineabdu(lLE,lin.getCell(5).asNumber().setScale(2, RoundingMode.HALF_UP), hoja5,fech,lin.getCellText(8),lin.getCellText(20),lin.getCellText(22),lin.getCellText(9));
								lLE++;
								break;
						}
					}
					String cuent=lin.getCellText(25);
					if(cuent.equals("0")) cuent=lin.getCellText(31);
					uti.linea(l,lin.getCell(5).asNumber().setScale(2, RoundingMode.HALF_UP).negate(), hoja,fech,lin.getCellText(8),lin.getCellText(20),cuent,lin.getCellText(9),asiento,comision+lin.getCellText(7).substring(lin.getCellText(7).length() - 1));
					l++;
					if(cuent.contains("487")) {
						switch(lin.getCellText(9)) {
							case "KT":
								uti.lineabdu(lKT,lin.getCell(5).asNumber().setScale(2, RoundingMode.HALF_UP).negate(), hoja1,fech,lin.getCellText(8),lin.getCellText(20),cuent,lin.getCellText(9));
								lKT++;
								break;
							case "PR":
								uti.lineabdu(lPR,lin.getCell(5).asNumber().setScale(2, RoundingMode.HALF_UP).negate(), hoja2,fech,lin.getCellText(8),lin.getCellText(20),cuent,lin.getCellText(9));
								lPR++;
								break;
							case "XN":
								uti.lineabdu(lXN,lin.getCell(5).asNumber().setScale(2, RoundingMode.HALF_UP).negate(), hoja3,fech,lin.getCellText(8),lin.getCellText(20),cuent,lin.getCellText(9));
								lXN++;
								break;
							case "KF":
								uti.lineabdu(lKF,lin.getCell(5).asNumber().setScale(2, RoundingMode.HALF_UP).negate(), hoja4,fech,lin.getCellText(8),lin.getCellText(20),cuent,lin.getCellText(9));
								lKF++;
								break;
							case "LE":
								uti.lineabdu(lLE,lin.getCell(5).asNumber().setScale(2, RoundingMode.HALF_UP).negate(), hoja5,fech,lin.getCellText(8),lin.getCellText(20),cuent,lin.getCellText(9));
								lLE++;
								break;
						}
					}
					break;
				case "FETA_DESFETA":
					BigDecimal devengo=new BigDecimal(lin.getCellText(17)).setScale(2, RoundingMode.HALF_UP);
					asiento="Periodificación FETA_DESFETA";
					comision="Periodificación "+lin.getCellText(20)+ " año ";
					
					//uti.linea(l,lin.getCell(17).asNumber(), hoja,fech,lin.getCellText(8),lin.getCellText(20),lin.getCellText(22),lin.getCellText(9));
					uti.linea(l,devengo, hoja,fech,lin.getCellText(8),lin.getCellText(20),lin.getCellText(22),lin.getCellText(9),asiento,comision+lin.getCellText(7).substring(lin.getCellText(7).length() - 1));
					l++;
					if(lin.getCellText(22).contains("487")) {
						switch(lin.getCellText(9)) {
							case "KT":
								uti.lineabdu(lKT,devengo.setScale(2, RoundingMode.HALF_UP), hoja1,fech,lin.getCellText(8),lin.getCellText(20),lin.getCellText(22),lin.getCellText(9));
								lKT++;
								break;
							case "PR":
								uti.lineabdu(lPR,devengo.setScale(2, RoundingMode.HALF_UP), hoja2,fech,lin.getCellText(8),lin.getCellText(20),lin.getCellText(22),lin.getCellText(9));
								lPR++;
								break;
							case "XN":
								uti.lineabdu(lXN,devengo.setScale(2, RoundingMode.HALF_UP), hoja3,fech,lin.getCellText(8),lin.getCellText(20),lin.getCellText(22),lin.getCellText(9));
								lXN++;
								break;
							case "KF":
								uti.lineabdu(lKF,devengo.setScale(2, RoundingMode.HALF_UP), hoja4,fech,lin.getCellText(8),lin.getCellText(20),lin.getCellText(22),lin.getCellText(9));
								lKF++;
								break;
							case "LE":
								uti.lineabdu(lLE,devengo.setScale(2, RoundingMode.HALF_UP), hoja5,fech,lin.getCellText(8),lin.getCellText(20),lin.getCellText(22),lin.getCellText(9));
								lLE++;
								break;
						}
					}
					uti.linea(l,devengo.negate(), hoja,fech,lin.getCellText(8),lin.getCellText(20),lin.getCellText(25),lin.getCellText(9),asiento,comision+lin.getCellText(7).substring(lin.getCellText(7).length() - 1));
					l++;
					if(lin.getCellText(25).contains("487")) {
						switch(lin.getCellText(9)) {
							case "KT":
								uti.lineabdu(lKT,devengo.setScale(2, RoundingMode.HALF_UP).negate(), hoja1,fech,lin.getCellText(8),lin.getCellText(20),lin.getCellText(25),lin.getCellText(9));
								lKT++;
								break;
							case "PR":
								uti.lineabdu(lPR,devengo.setScale(2, RoundingMode.HALF_UP).negate(), hoja2,fech,lin.getCellText(8),lin.getCellText(20),lin.getCellText(25),lin.getCellText(9));
								lPR++;
								break;
							case "XN":
								uti.lineabdu(lXN,devengo.setScale(2, RoundingMode.HALF_UP).negate(), hoja3,fech,lin.getCellText(8),lin.getCellText(20),lin.getCellText(25),lin.getCellText(9));
								lXN++;
								break;
							case "KF":
								uti.lineabdu(lKF,devengo.setScale(2, RoundingMode.HALF_UP).negate(), hoja4,fech,lin.getCellText(8),lin.getCellText(20),lin.getCellText(25),lin.getCellText(9));
								lKF++;
								break;
							case "LE":
								uti.lineabdu(lLE,devengo.setScale(2, RoundingMode.HALF_UP).negate(), hoja5,fech,lin.getCellText(8),lin.getCellText(20),lin.getCellText(25),lin.getCellText(9));
								lLE++;
								break;
						}
					}
					uti.linea(l,devengo, hoja,fech2,lin.getCellText(8),lin.getCellText(20),lin.getCellText(25),lin.getCellText(9),asiento,comision+lin.getCellText(7).substring(lin.getCellText(7).length() - 1));
					l++;
					uti.linea(l,devengo.negate(), hoja,fech2,lin.getCellText(8),lin.getCellText(20),lin.getCellText(22),lin.getCellText(9),asiento,comision+lin.getCellText(7).substring(lin.getCellText(7).length() - 1));
					l++;
					break;
				}
				
			}
			depurar(hoja1,hoja6);
			depurar(hoja2,hoja7);
			depurar(hoja3,hoja8);
			depurar(hoja4,hoja9);
			depurar(hoja5,hoja10);
			
			salida.finish();			
			outputStream.close();
			
            
            JOptionPane.showMessageDialog(this, "Fichero creado","Información", JOptionPane.INFORMATION_MESSAGE);
		} catch (Exception e) {
			JOptionPane.showMessageDialog(this, e.toString(),"Error salida", JOptionPane.ERROR_MESSAGE);	
		}

	}
	
	/**
	 * Método para introducir las fechas
	 */
	private void fechas() {
		JDateChooser jd1 = new JDateChooser();
		JDateChooser jd2 = new JDateChooser();
		String mensaje1 ="Elige fecha último día mes:\n";
		String mensaje2 ="Elige fecha primer día mes:\n";
		Object[] params = {mensaje1,jd1,mensaje2,jd2};
		if (JOptionPane.showConfirmDialog(this, params, "Fechas", JOptionPane.PLAIN_MESSAGE) == 0) { 
			if(jd1.getDate()!=null && jd2.getDate()!=null) {
				if(jd1.getDate().before(jd2.getDate())) {
					fech=jd1.getDate();
					fech2=jd2.getDate();
				} else {
					JOptionPane.showMessageDialog(this, "El primer día del mes ha de ser posterior al último día del mes","Error fechas", JOptionPane.ERROR_MESSAGE);
					fechas();
				}
			} else {
				JOptionPane.showMessageDialog(this, "Debe elegir ambas fechas","Error fechas", JOptionPane.ERROR_MESSAGE);
				fechas();
			}
		} else {
			fechas();
		}
	}

	/**
	 * Método de salida, donde cierra todos los libros y ficheros usados
	 */
	private void salida() {
		try {
			System.exit(0);
		} catch (Exception e) {
			JOptionPane.showMessageDialog(this, e.toString(),"Error salida", JOptionPane.ERROR_MESSAGE);
		}
	}

	public static void main(String[] args) {
		@SuppressWarnings("unused")
		Emgbdu eje=new Emgbdu();

	}
	
	private void guardarPrueba(List<String> prueba) {
		try {
			//final String nombreArchivo = "ComisionesRepercutidas.xlsx";
			String nombreArchivo =JOptionPane.showInputDialog("Nombre del archivo de salida: ");
			if(nombreArchivo==null) {
				nombreArchivo="Mierder.xlsx";
			} else {
				nombreArchivo+=".xlsx";
			}
			File directorioActual = new File(".");
			String ubicacion = directorioActual.getAbsolutePath();
			String ubicacionArchivoSalida = ubicacion.substring(0, ubicacion.length() - 1) + nombreArchivo;
			FileOutputStream outputStream;
			outputStream = new FileOutputStream(ubicacionArchivoSalida);
			@SuppressWarnings("resource")
			Workbook salida=new Workbook(outputStream,this.getTitle(),null);
			Worksheet hoja=salida.newWorksheet("EMG");
			
			
			for(int i=0;i<prueba.size();i++) {
				
				hoja.value(i, 0,prueba.get(i));			
			}
			
			
			salida.finish();			
			outputStream.close();
			
            
            JOptionPane.showMessageDialog(this, "Fichero creado","Información", JOptionPane.INFORMATION_MESSAGE);
		} catch (Exception e) {
			JOptionPane.showMessageDialog(this, e.toString(),"Error salida", JOptionPane.ERROR_MESSAGE);	
		}

	}
/*	private void guardarPrueba() {
		try {
			//final String nombreArchivo = "ComisionesRepercutidas.xlsx";
			String nombreArchivo =JOptionPane.showInputDialog("Nombre del archivo de salida: ");
			if(nombreArchivo==null) {
				nombreArchivo="Mierder.xlsx";
			} else {
				nombreArchivo+=".xlsx";
			}
			File directorioActual = new File(".");
			String ubicacion = directorioActual.getAbsolutePath();
			String ubicacionArchivoSalida = ubicacion.substring(0, ubicacion.length() - 1) + nombreArchivo;
			FileOutputStream outputStream;
			outputStream = new FileOutputStream(ubicacionArchivoSalida);
			@SuppressWarnings("resource")
			Workbook salida=new Workbook(outputStream,this.getTitle(),null);
			Worksheet hoja=salida.newWorksheet("EMG");
			uti.cabeceraEMGestion(hoja);
			Row lin;
			int l=1;
			String asiento,comision;
			
			for(int i=0;i<periodificaciones.size();i++) {
				
				lin=periodificaciones.get(i);
				
				//    abono-5  devengado-17 operacion-8 situacion-20 cuentaIngreso-22  cuentaResta-25  tipo-9
					asiento="Periodificación FETA/Asiento en firme";
					comision="Periodificación "+lin.getCellText(20)+ " año ";
					uti.linea(l,lin.getCell(5).asNumber().setScale(2, RoundingMode.HALF_UP), hoja,fech,lin.getCellText(8),lin.getCellText(20),lin.getCellText(22),lin.getCellText(9),asiento,comision+lin.getCellText(7).substring(lin.getCellText(7).length() - 1));
					l++;				
			}
			
			
			salida.finish();			
			outputStream.close();
			
            
            JOptionPane.showMessageDialog(this, "Fichero creado","Información", JOptionPane.INFORMATION_MESSAGE);
		} catch (Exception e) {
			JOptionPane.showMessageDialog(this, e.toString(),"Error salida", JOptionPane.ERROR_MESSAGE);	
		}

	}*/

}
