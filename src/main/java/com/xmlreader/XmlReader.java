package com.xmlreader;

import java.io.File;
import java.util.List;

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

public class XmlReader {

	public Workbook getAndReadXml(List<String> files) throws Exception {
		System.out.println("getAndReadXml");

		File xmlFile = new File(files.get(0));

		ColumnSetup cs = new ColumnSetup();
		// workbook = cs.initXls();
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

					contador++;
				}

			}
		}

		return workbook;
	}
	

	
}
