package com.dascalitas;

import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.w3c.dom.Document;
import org.w3c.dom.Element;
import org.w3c.dom.Node;
import org.w3c.dom.NodeList;
import org.xml.sax.SAXException;

import javax.xml.parsers.DocumentBuilder;
import javax.xml.parsers.DocumentBuilderFactory;
import javax.xml.parsers.ParserConfigurationException;
import javax.xml.xpath.XPath;
import javax.xml.xpath.XPathConstants;
import javax.xml.xpath.XPathExpressionException;
import javax.xml.xpath.XPathFactory;
import java.io.*;
import java.util.Calendar;
import java.util.Date;

public class CreatePOIfromXML {
    public static void main(String[] args) {
        Workbook wb = new HSSFWorkbook();
        CreationHelper createHelper = wb.getCreationHelper();

        try {
            File inputFile = new File("company.xml");
            DocumentBuilderFactory dbFactory = DocumentBuilderFactory.newInstance();
            DocumentBuilder dBuilder;

            dBuilder = dbFactory.newDocumentBuilder();

            Document doc = dBuilder.parse(inputFile);
            doc.getDocumentElement().normalize();

            XPath xPath =  XPathFactory.newInstance().newXPath();

            String expression = "//department";
            NodeList nodeList = (NodeList) xPath.compile(expression).evaluate(
                    doc, XPathConstants.NODESET);

            for (int i = 0; i < nodeList.getLength(); i++) {
                Node nNode = nodeList.item(i);

                if (nNode.getNodeType() == Node.ELEMENT_NODE) {
                    Element eElement = (Element) nNode;

                    Sheet sheet = wb.createSheet(eElement.getAttribute("depId") + " - " + eElement.getAttribute("name"));

                    Row row = sheet.createRow(0);
                    CellStyle cellStyle = wb.createCellStyle();
                    cellStyle.setAlignment(HorizontalAlignment.CENTER);
                    cellStyle.setVerticalAlignment(VerticalAlignment.BOTTOM);
                    HSSFFont font = (HSSFFont) wb.createFont();
                    font.setBold(true);
                    cellStyle.setFont(font);

                    Cell cell0 = row.createCell(0);
                    cell0.setCellValue("ID");
                    cell0.setCellStyle(cellStyle);
                    Cell cell1 = row.createCell(1);
                    cell1.setCellValue("first Name");
                    cell1.setCellStyle(cellStyle);
                    Cell cell2 = row.createCell(2);
                    cell2.setCellValue("last Name");
                    cell2.setCellStyle(cellStyle);
                    Cell cell3 = row.createCell(3);
                    cell3.setCellValue("Birth Date");
                    cell3.setCellStyle(cellStyle);
                    Cell cell4 = row.createCell(4);
                    cell4.setCellValue("Position");
                    cell4.setCellStyle(cellStyle);
                    Cell cell5 = row.createCell(5);
                    cell5.setCellValue("skills");
                    cell5.setCellStyle(cellStyle);
                    sheet.addMergedRegion(new CellRangeAddress(0, 0, 5, 7));

                    NodeList employeeList = eElement.getElementsByTagName("member");

                    for (int emp = 0; emp < employeeList.getLength(); emp++) {

                        Node empNode = employeeList.item(emp);
                        Row row2 = sheet.createRow(emp+1);
                        DataFormat format = wb.createDataFormat();
                        CellStyle cellStyle2 = wb.createCellStyle();
                        cellStyle2.setDataFormat(createHelper.createDataFormat().getFormat("mm/dd/yyyy"));

                        if (empNode.getNodeType() == Node.ELEMENT_NODE) {
                            Element empElement = (Element) empNode;

                            row2.createCell(0).setCellValue(empElement.getAttribute("memId"));
                            row2.createCell(1).setCellValue(empElement
                                    .getElementsByTagName("firstName")
                                    .item(0)
                                    .getTextContent());
                            row2.createCell(2).setCellValue(empElement
                                    .getElementsByTagName("lastName")
                                    .item(0)
                                    .getTextContent());
                            Cell cellN = row2.createCell(3);
                            cellN.setCellStyle(cellStyle2);
                            cellN.setCellValue(empElement
                                    .getElementsByTagName("birthDate")
                                    .item(0)
                                    .getTextContent());
                            row2.createCell(4).setCellValue(empElement
                                    .getElementsByTagName("position")
                                    .item(0)
                                    .getTextContent());

                            for (int skill = 0; skill < empElement.getElementsByTagName("skill").getLength(); skill++) {
                                row2.createCell(skill+5).setCellValue(empElement.getElementsByTagName("skill").item(skill).getTextContent());
                        }
                    }
                    }
                }
            }
        } catch (ParserConfigurationException e) {
            e.printStackTrace();
        } catch (SAXException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        } catch (XPathExpressionException e) {
            e.printStackTrace();
        }

        try (OutputStream fileOut = new FileOutputStream("company.xls")) {
            wb.write(fileOut);
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}
