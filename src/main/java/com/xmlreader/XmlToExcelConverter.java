package com.xmlreader;

import java.awt.Font;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;
import java.util.Optional;
import java.util.logging.Level;
import java.util.logging.Logger;

import javax.swing.ImageIcon;
import javax.swing.JButton;
import javax.swing.JFileChooser;
import javax.swing.JFrame;
import javax.swing.JLabel;
import javax.swing.JPanel;
import javax.swing.UIManager;
import javax.swing.filechooser.FileNameExtensionFilter;

import org.apache.commons.io.FilenameUtils;
import org.apache.poi.ss.usermodel.Workbook;

import com.formdev.flatlaf.FlatLightLaf;
import com.xmlreader.server.NTPService;

/**
 * Created by Filipe Ares (filipe.ares@jbaysolutions.com) -
 * http://blog.jbaysolutions.com Date: 17-08-2015.
 */
public class XmlToExcelConverter extends Constants implements ActionListener {
	private static Workbook workbook = null;
	private static int rowNum = 0;
	static JLabel jLabel;
	String path = null;
	CustomJFileChooser fileChooser;
	static JButton buttonConvert;
	static JButton buttonDir;
	static private final Logger LOGGER = Logger.getLogger("");
	private static JFrame jFrame = null;
	// private final static int REGIMEN_FISCAL_COLUMN = 0;
	// private final static int RFC_COLUMN = 1;
	// private final static int NOMBRE_COLUMN = 2;

	XmlToExcelConverter() {
	}

	public static void main(String[] args) {
		
		/*try {
            // Set System L&F
        UIManager.setLookAndFeel(
            UIManager.getSystemLookAndFeelClassName());
    } 
    catch (UnsupportedLookAndFeelException e) {
       // handle exception
    }
    catch (ClassNotFoundException e) {
       // handle exception
    }
    catch (InstantiationException e) {
       // handle exception
    }
    catch (IllegalAccessException e) {
       // handle exception
    }*/
		
		try {
		    UIManager.setLookAndFeel( new FlatLightLaf() );
		} catch( Exception ex ) {
		    System.err.println( "Failed to initialize LaF" );
		}

		ImageIcon img = new ImageIcon("E:\\UNIDAD X\\APPS\\git\\xml-to-excel\\2xml2excel.jpg");
		
		
		
    		System.out.println("Starting application");
		// This will be title for the frame
		jFrame = new JFrame("Convertidor de archivos XML a Excel");
		jFrame.setIconImage(img.getImage());
		// given width & height will set up the modal width & height
		jFrame.setSize(720, 350);
		jFrame.setLocationRelativeTo(null);
		jFrame.setVisible(true);
		jFrame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
		// creating object of the current class
		XmlToExcelConverter xml2Excel = new XmlToExcelConverter();
		buttonConvert = new JButton("Convertir");
		buttonConvert.setEnabled(false);
		
		
		 NTPService servicioNTP = new NTPService();
		 
		    
		 
	        LOGGER.log(Level.INFO, "Obteniendo fecha desde el servidor NTP");
	        
	        Date dateFromServer = servicioNTP.getNTPDate();
	        
	        System.out.println("");
	        System.out.println("");
	        System.out.println("Fecha-Hora NTP: " + dateFromServer);
		
		
		//Date fechaActual = new Date(System.currentTimeMillis());
		String fecha = "2024-04-30";
		SimpleDateFormat date = new SimpleDateFormat("yyyy-MM-dd");
		Date fecha2 = null;
		try {
			fecha2 = date.parse(fecha);
		} catch (ParseException e) {
			e.printStackTrace();
		}
		buttonDir = new JButton("Abrir archivo o folder");
		jLabel = new JLabel("Selecciona un directorio o un archivo");
		//jLabel.setFont(new Font("Segoe UI", Font.BOLD, 14));

		if(dateFromServer.before(fecha2)){			
			buttonDir.setEnabled(true);
		}else {
			buttonDir.setEnabled(false);
			jLabel.setText("Licencia caducada");
		}
			
		buttonDir.addActionListener(xml2Excel);
		buttonConvert.addActionListener(xml2Excel);
		JPanel jP = new JPanel();
		jP.add(buttonConvert);
		jP.add(buttonDir);		
		jP.add(jLabel);
		jFrame.add(jP);		
		jFrame.setVisible(true);
		
	}

	public void actionPerformed(ActionEvent ae) {

		String flag = ae.getActionCommand();
		if (flag.equals("Abrir archivo o folder")) {

			fileChooser = new CustomJFileChooser();			
			
			if (Optional.ofNullable(path).isPresent()) {
				fileChooser.setCurrentDirectory(new File(path));
			} else {
				fileChooser.setCurrentDirectory(new File("."));
			}

			FileNameExtensionFilter filter = new FileNameExtensionFilter("XML File", "xml");

			fileChooser.setFileFilter(filter);
			fileChooser.setDialogTitle("Selecciona el archivo XML a convertir");
			fileChooser.setFileSelectionMode(JFileChooser.FILES_AND_DIRECTORIES);

			int dialogVal = fileChooser.showOpenDialog(jFrame);
			if (dialogVal == JFileChooser.APPROVE_OPTION) {
				jLabel.setText(fileChooser.getSelectedFile().getAbsolutePath());

				List<String> files = new ArrayList<>();

				path = fileChooser.getSelectedFile().getAbsolutePath();
				files.add(fileChooser.getSelectedFile().getAbsolutePath());

				XmlReader xmlReader = new XmlReader();
				try {
					workbook = xmlReader.getAndReadXml(files);
				} catch (Exception e) {
					e.printStackTrace();
				}

				buttonConvert.setEnabled(true);

			} else {
				jLabel.setText("¡Selección de archivo cancelada!");
			}
		}
		if (flag.equals("Convertir")) {

			String pathAndFileName = fileChooser.getSelectedFile().getAbsolutePath();
			String file = FilenameUtils.getBaseName(pathAndFileName);
			String folder = FilenameUtils.getFullPath(pathAndFileName);

			String saveFile = folder + file + ".xlsx";

			if (Optional.ofNullable(saveFile).isPresent()) {
				fileChooser.setCurrentDirectory(new File(folder));

			} else {
				fileChooser.setCurrentDirectory(new File("."));
			}

			FileNameExtensionFilter filter = new FileNameExtensionFilter("Archivo XLSX de Excel", "xlsx");

			fileChooser.setSelectedFile(new File(file + ".xlsx"));
			fileChooser.setFileFilter(filter);
			fileChooser.setDialogTitle("Guardar como archivo de excel");
			int dialogVal = fileChooser.showSaveDialog(jFrame);
			if (dialogVal == JFileChooser.APPROVE_OPTION) {

				// jLabel.setText(fileChooser.getSelectedFile().getAbsolutePath());
				jLabel.setText("Conversión completa");

				String pathAndFileName2Save = fileChooser.getSelectedFile().getAbsolutePath();
				String file2Save = FilenameUtils.getBaseName(pathAndFileName2Save);
				String pathFile = FilenameUtils.getFullPath(pathAndFileName2Save);

				String save2File = pathFile + file2Save + ".xlsx";

				FileOutputStream fileOut;
				try {
					fileOut = new FileOutputStream(save2File);
					workbook.write(fileOut);
					workbook.close();
					fileOut.close();
				} catch (FileNotFoundException e) {
					e.printStackTrace();
				} catch (IOException e) {
					e.printStackTrace();
				}
				buttonConvert.setEnabled(false);

			} else {
				jLabel.setText("¡Conversión cancelada!");
			}
		}
	}

	

}
