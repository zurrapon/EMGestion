package emgestion.bdu;

import java.io.File;
import java.io.IOException;
import java.util.Arrays;
import java.util.List;
import java.util.stream.Collectors;
import java.util.stream.Stream;

import javax.swing.JDialog;
import javax.swing.JFrame;
import javax.swing.JOptionPane;
import javax.swing.JPanel;
import javax.swing.JProgressBar;
import javax.swing.SwingWorker;

import org.dhatim.fastexcel.reader.CellType;
import org.dhatim.fastexcel.reader.ReadableWorkbook;
import org.dhatim.fastexcel.reader.Row;
import org.dhatim.fastexcel.reader.Sheet;

public class CargadorFicheros extends SwingWorker<Integer,Integer>{

    private final JProgressBar progreso;
    private ReadableWorkbook libro;
    private Utilidades uti=new Utilidades();
    private List<Row> periodificaciones;
    private List<Integer> anos;
    private String txt;//= new ArrayList<Row>();
    private JDialog ven;
    private JPanel fondo;
    private int a;
    JFrame ventana;

    public CargadorFicheros(JFrame ve,JDialog vent,JPanel fon,JProgressBar barra,File f1,List<Row> lista,String texto,List<Integer> ans) {
        ven=vent;
        fondo=fon;
    	progreso = barra;
        periodificaciones=lista;
        txt=texto;
        anos=ans;
        ventana=ve;
        
        if(txt.contains("Comisiones")) {
        	a=0;
        } else {
        	a=1;
        }
        try {				
			libro= new ReadableWorkbook(f1);
		    
		} catch (IOException e) {
			JOptionPane.showMessageDialog(null, e.toString(),"Error salida", JOptionPane.ERROR_MESSAGE);
		}
        fondo.updateUI();
        
    }

    @Override
    public Integer doInBackground() throws Exception {
    	List<Sheet> hojis=libro.getSheets().collect(Collectors.toList());
		int aumento=100/hojis.size();
		for(int i=0;i<hojis.size();i++) {
			Sheet pag=hojis.get(i);
			publish(aumento*(i+1));
			if(pag.getName().contains(txt)) {
				anos.set(a,anos.get(a)+1);
				List<Row> linis=pag.openStream().collect(Collectors.toList());
				int y=uti.primeraFila(linis);
				for( int x=y;x<linis.size() && x>-1;x++) {
					Row lin=linis.get(x);
					if(lin.getCell(0)!=null) {
						if(lin.getCell(0).getType()!= CellType.EMPTY) {
							periodificaciones.add(lin);
						}
					}
					
				}
			}
		}
		libro.close();
        return 0;
    }

    @Override
    protected void done() {
    	publish(100);
    	ven.dispose();
    }
 
    @Override
    protected void process(@SuppressWarnings("rawtypes") List chunks) {
        //Actualizando la barra de progreso. Datos del publish.
        for(int s=0;s<chunks.size();s++) {
        	progreso.setValue((int) chunks.get(s));
        }
    }
}
