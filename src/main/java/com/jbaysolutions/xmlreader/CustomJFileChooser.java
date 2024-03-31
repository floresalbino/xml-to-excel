package com.jbaysolutions.xmlreader;

import java.io.File;

import javax.swing.JFileChooser;
import javax.swing.JOptionPane;

public class CustomJFileChooser extends JFileChooser{
	
	
	 /**
	 * 
	 */
	private static final long serialVersionUID = 1L;

	@Override
	    public void approveSelection(){
	        File f = getSelectedFile();
	        if(f.exists() && getDialogType() == SAVE_DIALOG){
	            int result = JOptionPane.showConfirmDialog(this,"Este nombre de archivo ya existe, ¿sobrescribir archivo?","Archivo existente",JOptionPane.YES_NO_OPTION);
	            switch(result){
	                case JOptionPane.YES_OPTION:
	                    super.approveSelection();
	                    return;
	                case JOptionPane.NO_OPTION:
	                    return;
	                case JOptionPane.CLOSED_OPTION:
	                    return;
	                case JOptionPane.CANCEL_OPTION:
	                    cancelSelection();
	                    return;
	            }
	        }
	        super.approveSelection();
	    }        
	
}
