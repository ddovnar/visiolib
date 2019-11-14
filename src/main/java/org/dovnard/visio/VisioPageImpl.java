package org.dovnard.visio;

import org.w3c.dom.Document;
import org.w3c.dom.Element;

import javax.xml.parsers.DocumentBuilder;
import javax.xml.parsers.DocumentBuilderFactory;
import javax.xml.parsers.ParserConfigurationException;
import javax.xml.transform.OutputKeys;
import javax.xml.transform.Transformer;
import javax.xml.transform.TransformerException;
import javax.xml.transform.TransformerFactory;
import javax.xml.transform.dom.DOMSource;
import javax.xml.transform.stream.StreamResult;
import java.io.File;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

/**
 * Visio project
 * Author: Dmitriy Dovnar
 * Date: 06.11.19
 * Time: 23:16
 */
public class VisioPageImpl implements VisioPage, XmlExportable {
    private final String name;
    private int id;
    private int rel_id;

    private List<Shape> shapes;
    private List<Connect> connects;

    public VisioPageImpl(String name) {
        this.name = name;
        this.shapes = new ArrayList<>();
        this.connects = new ArrayList<>();
    }

    public String getName() {
        return name;
    }

    @Override
    public Element getXmlElement(Document document) {
        Element root = document.createElement("Page");
        root.setAttribute("ID", Integer.toString(id));
        root.setAttribute("Name", name);
        root.setAttribute("NameU", name);
        root.setAttribute("IsCustomNameU", "1");
        root.setAttribute("ViewScale", "0.82");
        root.setAttribute("ViewCenterX", "4.1275082550165");
        root.setAttribute("ViewCenterY", "8.5852171704343");

        Element pageSheet = document.createElement("PageSheet");
        pageSheet.setAttribute("LineStyle", "0");
        pageSheet.setAttribute("FillStyle", "0");
        pageSheet.setAttribute("TextStyle", "0");

        Element cel1 = document.createElement("Cell");
        cel1.setAttribute("N", "PageWidth");
        cel1.setAttribute("V", "8.26771653543307");
        pageSheet.appendChild(cel1);

        Element cel2 = document.createElement("Cell");
        cel2.setAttribute("N", "PageHeight");
        cel2.setAttribute("V", "11.69291338582677");
        pageSheet.appendChild(cel2);

        Element cel3 = document.createElement("Cell");
        cel3.setAttribute("N", "ShdwOffsetX");
        cel3.setAttribute("V", "0.1181102362204724");
        pageSheet.appendChild(cel3);

        Element cel4 = document.createElement("Cell");
        cel4.setAttribute("N", "ShdwOffsetY");
        cel4.setAttribute("V", "-0.1181102362204724");
        pageSheet.appendChild(cel4);

        Element cel5 = document.createElement("Cell");
        cel5.setAttribute("N", "PageScale");
        cel5.setAttribute("V", "0.03937007874015748");
        cel5.setAttribute("U", "IN_F");
        pageSheet.appendChild(cel5);

        Element cel6 = document.createElement("Cell");
        cel6.setAttribute("N", "DrawingScale");
        cel6.setAttribute("V", "0.03937007874015748");
        cel6.setAttribute("U", "IN_F");
        pageSheet.appendChild(cel6);

        Element cel7 = document.createElement("Cell");
        cel7.setAttribute("N", "DrawingSizeType");
        cel7.setAttribute("V", "0");
        pageSheet.appendChild(cel7);

        Element cel8 = document.createElement("Cell");
        cel8.setAttribute("N", "DrawingScaleType");
        cel8.setAttribute("V", "0");
        pageSheet.appendChild(cel8);

        Element cel9 = document.createElement("Cell");
        cel9.setAttribute("N", "InhibitSnap");
        cel9.setAttribute("V", "0");
        pageSheet.appendChild(cel9);

        Element cel10 = document.createElement("Cell");
        cel10.setAttribute("N", "PageLockReplace");
        cel10.setAttribute("V", "0");
        cel10.setAttribute("U", "BOOL");
        pageSheet.appendChild(cel10);

        Element cel11 = document.createElement("Cell");
        cel11.setAttribute("N", "PageLockDuplicate");
        cel11.setAttribute("V", "0");
        cel11.setAttribute("U", "BOOL");
        pageSheet.appendChild(cel11);

        Element cel12 = document.createElement("Cell");
        cel12.setAttribute("N", "UIVisibility");
        cel12.setAttribute("V", "0");
        pageSheet.appendChild(cel12);

        Element cel13 = document.createElement("Cell");
        cel13.setAttribute("N", "ShdwType");
        cel13.setAttribute("V", "0");
        pageSheet.appendChild(cel13);

        Element cel14 = document.createElement("Cell");
        cel14.setAttribute("N", "ShdwObliqueAngle");
        cel14.setAttribute("V", "0");
        pageSheet.appendChild(cel14);

        Element cel15 = document.createElement("Cell");
        cel15.setAttribute("N", "ShdwScaleFactor");
        cel15.setAttribute("V", "1");
        pageSheet.appendChild(cel15);

        Element cel16 = document.createElement("Cell");
        cel16.setAttribute("N", "DrawingResizeType");
        cel16.setAttribute("V", "1");
        pageSheet.appendChild(cel16);

        Element cel17 = document.createElement("Cell");
        cel17.setAttribute("N", "PageShapeSplit");
        cel17.setAttribute("V", "1");
        pageSheet.appendChild(cel17);

        root.appendChild(pageSheet);

        Element rel = document.createElement("Rel");
        rel.setAttribute("r:id", getRelId());
        root.appendChild(rel);

        return root;
    }

    @Override
    public void setId(int id) {
        this.id = id;
    }

    @Override
    public void setRelId(int relId) {
        this.rel_id = relId;
    }

    @Override
    public int getId() {
        return id;
    }

    @Override
    public String getRelId() {
        return "rId" + rel_id;
    }

    @Override
    public String getShortFileName() {
        return "page" + id + ".xml";
    }

    @Override
    public Shape addShape(ShapeType shapeType, String name, double xpos, double ypos) {
        Shape newShape = ShapeFactory.createShape(shapeType, name, xpos, ypos);
        newShape.setId(shapes.size() + 1);
        shapes.add(newShape);
        return newShape;
    }

    @Override
    public Shape addShape(ShapeType shapeType, String name, String title, double xpos, double ypos) {
        Shape newShape = addShape(shapeType, name, xpos, ypos);
        newShape.setTitle(title);
        return newShape;
    }

    @Override
    public void addConnectorBottomTop(Shape source, Shape target) {
        ConnectorShape conn = new ConnectorShape();
        conn.setName("conn1");
        conn.setId(shapes.size() + 1);

        conn.setPinX((source.getXpos() + target.getXpos()) / 2);
        conn.setPinY((source.getYpos() + target.getYpos()) / 2);
        conn.setBeginX(source.getXpos());
        conn.setBeginY(source.getYpos() - (source.getHeight() / 2));

        conn.setEndX(target.getXpos());
        conn.setEndY(target.getYpos() + (target.getHeight() / 2));
//        conn.setBeginX(4.0);
//        conn.setBeginY(9.5);
//        conn.setEndX(1.0);
//        conn.setEndY(8.5);
//        conn.setLocPinX(-2.5);
        conn.setLocPinX(-((source.getXpos() - target.getXpos()) / 2));
//        conn.setLocPinY(-0.5);
        double dy = -(((source.getYpos() - (source.getHeight() / 2)) - (target.getYpos() + (target.getHeight() / 2)))) / 2;
        conn.setLocPinY(dy);

        conn.addDrawCmd("MoveTo", 0.0, 0.0);
        conn.addDrawCmd("LineTo", 0.0, dy);
//        conn.addDrawCmd("LineTo", -3.0, -0.5);
        conn.addDrawCmd("LineTo", -(source.getXpos() - target.getXpos()),dy);
        conn.addDrawCmd("LineTo", -(source.getXpos() - target.getXpos()), dy + dy);
//        conn.addDrawCmd("LineTo", -3.0, -1.0);

        shapes.add(conn);

        Connect sourceConnection = new Connect();
        sourceConnection.setFromSheet(Integer.toString(source.getId()));
        sourceConnection.setFromCell("EndX");
        sourceConnection.setFromPart("12");
        sourceConnection.setToSheet(Integer.toString(target.getId()));
        sourceConnection.setToCell("Connections.X3");
        sourceConnection.setToPart("102");
        connects.add(sourceConnection);

        Connect targetConnection = new Connect();
        targetConnection.setFromSheet(Integer.toString(source.getId()));
        targetConnection.setFromCell("BeginX");
        targetConnection.setFromPart("9");
        targetConnection.setToSheet(Integer.toString(target.getId()));
        targetConnection.setToCell("Connections.X1");
        targetConnection.setToPart("100");
        connects.add(targetConnection);
//        Connect c1 = new Connect();
//        c1.setFromSheet("3");
//        c1.setFromCell("EndX");
//        c1.setFromPart("12");
//        c1.setToSheet("2");
//        c1.setToCell("Connections.X3");
//        c1.setToPart("102");
//
//        connects.add(c1);
//
//        Connect c2 = new Connect();
//        c2.setFromSheet("3");
//        c2.setFromCell("BeginX");
//        c2.setFromPart("9");
//        c2.setToSheet("1");
//        c2.setToCell("Connections.X1");
//        c2.setToPart("100");
//
//        connects.add(c2);
    }

    @Override
    public Document toXml() throws ParserConfigurationException {
        DocumentBuilderFactory docFactory = DocumentBuilderFactory.newInstance();

        DocumentBuilder docBuilder = docFactory.newDocumentBuilder();
        Document document = docBuilder.newDocument();
        document.setXmlStandalone(true);

        Element root = document.createElement("PageContents");
        root.setAttribute("xmlns", "http://schemas.microsoft.com/office/visio/2012/main");
        root.setAttribute("xmlns:r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
        root.setAttribute("xml:space", "preserve");
        document.appendChild(root);

        if (shapes.size() > 0) {
            Element shapesEl = document.createElement("Shapes");
            root.appendChild(shapesEl);

            Iterator<Shape> shapeIter = shapes.iterator();
            while (shapeIter.hasNext()) {
                Shape shape = shapeIter.next();
                shapesEl.appendChild(shape.getXmlElement(document));
            }
        }

        if (connects.size() > 0) {
            Element connectsE = document.createElement("Connects");
            Iterator<Connect> conIter = connects.iterator();
            while (conIter.hasNext()) {
                Connect con = conIter.next();
                connectsE.appendChild(con.getXmlElement(document));
            }
            root.appendChild(connectsE);
        }

        return document;
    }

    @Override
    public String getFileName() {
        return "/visio/pages/page" + id + ".xml";
    }

    @Override
    public void writeToFile(String path) throws IOException, TransformerException, ParserConfigurationException {
        XmlFileWriter.writeToFile(path, getFileName(), toXml());
    }

    @Override
    public void writeToFile(String path, String filename) throws IOException, TransformerException, ParserConfigurationException {
        XmlFileWriter.writeToFile(path, filename, toXml());
    }
}
