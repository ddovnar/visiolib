package org.dovnard.visio;

import org.junit.Test;
import org.junit.Assert;
import org.w3c.dom.Document;
import org.w3c.dom.Element;

import javax.xml.parsers.DocumentBuilder;
import javax.xml.parsers.DocumentBuilderFactory;
import javax.xml.parsers.ParserConfigurationException;
import javax.xml.transform.Transformer;
import javax.xml.transform.TransformerConfigurationException;
import javax.xml.transform.TransformerException;
import javax.xml.transform.TransformerFactory;
import javax.xml.transform.dom.DOMSource;
import javax.xml.transform.stream.StreamResult;
import java.text.SimpleDateFormat;
import java.util.Date;

/**
 * Visio project
 * Author: Dmitriy Dovnar
 * Date: 06.11.19
 * Time: 23:08
 */
public class VisioDocumentTest {
    @Test
    public void test1() {
        VisioDocument visioDocument = null;
        try {
            visioDocument = VisioDocumentBuilder.newDocument();

            //visioDocument.create();
            VisioPage page = visioDocument.createPage("Page1");

            //VisioPage page1 = visioDocument.createPage("Page2");

            visioDocument.getPageList().stream().forEach(pageItem -> {
                System.out.println("Page: " + pageItem.getName());
            });
            Assert.assertTrue("Ok", visioDocument.getPageList().size() == 1);

//            ShapeDef[][] matrix = new ShapeDef[3][4];
//            matrix[0][0] = new ShapeDef("shape1", "Shape[0,0]", 1.0, 10.0);
//            matrix[0][1] = new ShapeDef("shape2", "Shape[0,1]", 3.0, 10.0);
//            matrix[0][2] = new ShapeDef("shape3", "Shape[0,2]", 5.0, 10.0);
//            matrix[0][3] = new ShapeDef("shape33", "Shape[0,3]", 7.0, 10.0);
//
//            matrix[1][0] = new ShapeDef("shape4", "Shape[1,0]", 1.0, 9.0);
//            matrix[1][1] = new ShapeDef("shape5", "Shape[1,1]", 3.0, 9.0);
//            matrix[1][2] = new ShapeDef("shape6", "Shape[1,2]", 5.0, 9.0);
//            matrix[1][3] = new ShapeDef("shape66", "Shape[1,3]", 7.0, 9.0);
//
//            matrix[2][0] = new ShapeDef("shape7", "Shape[2,0]", 1.0, 8.0);
//            matrix[2][1] = new ShapeDef("shape8", "Shape[2,1]", 3.0, 8.0);
//            matrix[2][2] = new ShapeDef("shape9", "Shape[2,2]", 5.0, 8.0);
//            matrix[2][3] = new ShapeDef("shape99", "Shape[2,3]", 7.0, 8.0);
//
//            for (int i = 0; i < matrix.length; i++) {
//                for (int j = 0; j < matrix[i].length; j++) {
//                    ShapeDef def = matrix[i][j];
//                    Shape shape = page.addShape(ShapeType.RECTANGLE, def.name, def.title, def.xpos, def.ypos);
//                    shape.setHeight(def.height);
//                }
//            }
            Shape rec1 = page.addShape(ShapeType.RECTANGLE, "test1", "root", 4.0, 10.0);
            rec1.setPrintDebugInfo(false);
            //rec1.setWidth(0.5);
            //rec1.setHeight(0.5);

            System.out.println("Width:" + rec1.getWidth());

            Shape rec2 = page.addShape(ShapeType.RECTANGLE, "test2", "Child1", 1.0, 8.0);
            rec2.setPrintDebugInfo(false);

            page.addConnectorBottomTop(rec1, rec2);
            //rec2.setHeight(0.5);
//
            Shape rec3 = page.addShape(ShapeType.RECTANGLE, "test3", "Child2", 3.0, 8.0);
            rec3.setPrintDebugInfo(false);
            page.addConnectorBottomTop(rec1, rec3);

            Shape rec4 = page.addShape(ShapeType.RECTANGLE, "test4", "Child3", 5.0, 8.0);
            rec4.setPrintDebugInfo(false);
            page.addConnectorBottomTop(rec1, rec4);

            Shape rec5 = page.addShape(ShapeType.RECTANGLE, "test5", "Child4", 7.0, 8.0);
            rec5.setPrintDebugInfo(false);
            page.addConnectorBottomTop(rec1, rec5);
//            rec3.setHeight(0.5);
            // visioDocument.writeToFile("test_file");

            visioDocument.createVisioFile(getUniqFileName("doc_"));
        } catch (Exception e) {
            e.printStackTrace();
            Assert.fail();
        }
    }

    private String getUniqFileName(String fname) {
        String ftime = new SimpleDateFormat("yyyyMMddHHmm'.vsdx'").format(new Date());

        return fname + ftime;
    }

    public void test2() {
        DocumentBuilderFactory docFactory = DocumentBuilderFactory.newInstance();
        try {
            DocumentBuilder docBuilder = docFactory.newDocumentBuilder();
            Document document = docBuilder.newDocument();

            Element root = document.createElement("root");
            document.appendChild(root);

            DOMSource source = new DOMSource(document);
            StreamResult result = new StreamResult(System.out);

            TransformerFactory transformerFactory = TransformerFactory.newInstance();
            Transformer transformer = transformerFactory.newTransformer();
            transformer.transform(source, result);

        } catch (ParserConfigurationException e) {
            e.printStackTrace();
        } catch (TransformerConfigurationException e) {
            e.printStackTrace();
        } catch (TransformerException e) {
            e.printStackTrace();
        }
    }

    class ShapeDef {
        String name;
        String title;
        double xpos;
        double ypos;
        double height;
        public ShapeDef(String name, String title, double xpos, double ypos) {
            this.name = name;
            this.title = title;
            this.xpos = xpos;
            this.ypos = ypos;
            this.height = 0.5;
        }
    }
}
