package com.jbaysolutions.xmlreader;

import java.io.File;
import java.io.FileOutputStream;

import javax.xml.parsers.DocumentBuilder;
import javax.xml.parsers.DocumentBuilderFactory;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.w3c.dom.Document;
import org.w3c.dom.Element;
import org.w3c.dom.NamedNodeMap;
import org.w3c.dom.Node;
import org.w3c.dom.NodeList;

/**
 * Created by Filipe Ares (filipe.ares@jbaysolutions.com) -
 * http://blog.jbaysolutions.com Date: 17-08-2015.
 */
public class XmlToExcelConverter extends Constants {
	private static Workbook workbook;
	private static int rowNum = 0;

	// private final static int REGIMEN_FISCAL_COLUMN = 0;
	// private final static int RFC_COLUMN = 1;
	// private final static int NOMBRE_COLUMN = 2;

	public static void main(String[] args) throws Exception {

		System.out.println("Starting application");
		getAndReadXml();
	}

	/**
	 *
	 * Downloads a XML file, reads the substance and product values and then writes
	 * them to rows on an excel file.
	 *
	 * @throws Exception
	 */
	private static void getAndReadXml() throws Exception {
		System.out.println("getAndReadXml");

		/*
		 * File xmlFile = File.createTempFile("substances", "tmp"); String xmlFileUrl =
		 * "http://ec.europa.eu/food/plant/pesticides/eu-pesticides-database/public/?event=Execute.DownLoadXML&id=1";
		 * URL url = new URL(xmlFileUrl); System.out.println("downloading file from " +
		 * xmlFileUrl + " ..."); FileUtils.copyURLToFile(url, xmlFile);
		 * System.out.println("downloading finished, parsing...");
		 */
		File xmlFile = new File("C:/Temp/0c334547-ad15-4b4f-9d81-be840b6c4e81.xml");

		/*
		 * If you have the xml file locally, replace the above code by the following
		 * line: File xmlFile = new File("C:/Temp/Publication1.xml");
		 */
		ColumnSetup cs = new ColumnSetup();
		//workbook = cs.initXls();
		Workbook workbook = new XSSFWorkbook();
		Row row = null;
		Row row2 = null;
		Sheet sheet = workbook.createSheet();
		row = sheet.createRow(0);
		row2 = sheet.createRow(1);
	

		DocumentBuilderFactory dbFactory = DocumentBuilderFactory.newInstance();
		DocumentBuilder dBuilder = dbFactory.newDocumentBuilder();
		Document doc = dBuilder.parse(xmlFile);

		 
		    NodeList nodeList = doc.getElementsByTagName("*");
		    int contador = 0;
		    for (int i = 0; i < nodeList.getLength(); i++) {
		        Node node = nodeList.item(i);
		        if (node.getNodeType() == Node.ELEMENT_NODE) {

		        	System.out.println("/***************/" + node.getNodeName() + "/***************/");	
		            
		            Element element = (Element) node;
		            NamedNodeMap namedNodeMap = element.getAttributes();
		            for (int j = 0; j < namedNodeMap.getLength(); j++) {
		            	
						Cell cell = row.createCell(contador);
						cell.setCellValue(node.getNodeName() + "--" + element.getAttributes().item(j).getNodeName());				
						
						cell = row2.createCell(contador);
						cell.setCellValue(element.getAttributes().item(j).getNodeValue());
		            	
		            	System.out.println(element.getAttributes().item(j).getNodeName());
		            	System.out.println(element.getAttributes().item(j).getNodeValue());
		            	System.out.println(contador);
		            	
		            	contador ++;
		            }
		            
		            
		            
		        }
		    }
		
		
		/*
		NodeList nList = doc.getElementsByTagName("cfdi:Emisor");
		for (int i = 0; i < nList.getLength(); i++) {
			System.out.println("Processing element " + (i + 1) + "/" + nList.getLength());
			Node node = nList.item(i);
			if (node.getNodeType() == Node.ELEMENT_NODE) {
				Element element = (Element) node;
				String column = null;//element.getAttributes().getNamedItem(REGIMEN_FISCAL).getNodeValue();
				
				//column.
				
				String rfc = element.getAttributes().getNamedItem(RFC_EMISOR).getNodeValue();
				String nombre = element.getAttributes().getNamedItem(NOMBRE_EMISOR).getNodeValue();

				row = sheet.createRow(rowNum);
				Cell cell = row.createCell(i);
				cell.setCellValue(element.getChildNodes().item(i).getNodeName() + EMISOR);				
				
				row = sheet.createRow(++rowNum);
				cell = row.createCell(i);
				cell.setCellValue(column);
/*
				cell = row.createCell(1);
				cell.setCellValue(rfc);

				cell = row.createCell(2);
				cell.setCellValue(nombre);
*/
				// .getElementsByTagName("RegimenFiscal").item(0).getTextContent();
				// String entryForce =
				// element.getElementsByTagName("entry_force").item(0).getTextContent();
				// String directive =
				// element.getElementsByTagName("directive").item(0).getTextContent();

				/*
				 * NodeList prods = element.getElementsByTagName("Product"); for (int j = 0; j <
				 * prods.getLength(); j++) { Node prod = prods.item(j); if (prod.getNodeType()
				 * == Node.ELEMENT_NODE) { Element product = (Element) prod; String prodName =
				 * product.getElementsByTagName("Product_name").item(0).getTextContent(); String
				 * prodCode =
				 * product.getElementsByTagName("Product_code").item(0).getTextContent(); String
				 * lmr = product.getElementsByTagName("MRL").item(0).getTextContent(); String
				 * applicationDate =
				 * product.getElementsByTagName("ApplicationDate").item(0).getTextContent();
				 * 
				 * Row row = sheet.createRow(rowNum++); Cell cell =
				 * row.createCell(SUBSTANCE_NAME_COLUMN); cell.setCellValue(regimenFiscal);
				 * 
				 * cell = row.createCell(SUBSTANCE_ENTRY_FORCE_COLUMN); cell.setCellValue(rfc);
				 * 
				 * cell = row.createCell(SUBSTANCE_DIRECTIVE_COLUMN); cell.setCellValue(nombre);
				 * /* cell = row.createCell(PRODUCT_NAME_COLUMN); cell.setCellValue(prodName);
				 * 
				 * cell = row.createCell(PRODUCT_CODE_COLUMN); cell.setCellValue(prodCode);
				 * 
				 * cell = row.createCell(PRODUCT_MRL_COLUMN); cell.setCellValue(lmr);
				 * 
				 * cell = row.createCell(APPLICATION_DATE_COLUMN);
				 * cell.setCellValue(applicationDate);
				 * 
				 * } }
				 
			}
		}

		nList = doc.getElementsByTagName("cfdi:Receptor");
		for (int i = 0; i < nList.getLength(); i++) {
			System.out.println("Processing element " + (i + 1) + "/" + nList.getLength());
			Node node = nList.item(i);
			if (node.getNodeType() == Node.ELEMENT_NODE) {
				Element element = (Element) node;
				String rfcReceptor = element.getAttributes().getNamedItem(RFC_RECEPTOR).getNodeValue();
				String nombreReceptor = element.getAttributes().getNamedItem(NOMBRE_RECEPTOR).getNodeValue();
				String usoCFDI = element.getAttributes().getNamedItem(USO_CFDI).getNodeValue();

				Cell cell = row.createCell(3);
				cell.setCellValue(rfcReceptor);

				cell = row.createCell(4);
				cell.setCellValue(nombreReceptor);

				cell = row.createCell(5);
				cell.setCellValue(usoCFDI);
			}
		}
*/
		FileOutputStream fileOut = new FileOutputStream("C:/Temp/Excel-Out.xlsx");
		workbook.write(fileOut);
		workbook.close();
		fileOut.close();

		/*
		 * if (xmlFile.exists()) { System.out.println("delete file-> " +
		 * xmlFile.getAbsolutePath()); if (!xmlFile.delete()) {
		 * System.out.println("file '" + xmlFile.getAbsolutePath() +
		 * "' was not deleted!"); } }
		 */

		//System.out.println("getAndReadXml finished, processed ");
	}

}
