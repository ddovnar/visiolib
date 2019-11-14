package org.dovnard.visio;

import org.w3c.dom.Document;
import org.w3c.dom.Element;

import javax.swing.text.AbstractDocument;

/**
 * User: DovnarDmitriy
 * Date: 08.11.19
 * Time: 16:47
 */
public abstract class Shape implements XmlContent {

    private int id;
    private double xpos;
    private double ypos;
    private String name;
    private String title;
    private double width;
    private double height;
    private boolean printDebugInfo;

    public Shape() {}
    public Shape(String name, double xpos, double ypos) {
        this.name = name;
        this.xpos = xpos;
        this.ypos = ypos;
    }

    public Shape(String name, String title, double xpos, double ypos) {
        this.name = name;
        this.xpos = xpos;
        this.ypos = ypos;
        this.title = title;
    }

    public double getWidth() {
        return width;
    }

    public void setWidth(double width) {
        this.width = width;
    }

    public double getHeight() {
        return height;
    }

    public void setHeight(double height) {
        this.height = height;
    }

    public String getTitle() {
        if (!printDebugInfo)
            return title;
        return title + getDebugInfo();
    }

    public void setPrintDebugInfo(boolean printDebugInfo) {
        this.printDebugInfo = printDebugInfo;
    }

    public void setTitle(String title) {
        this.title = title;
    }

    public String getName() {
        return name;
    }

    public void setName(String name) {
        this.name = name;
    }

    public int getId() {
        return id;
    }

    public void setId(int id) {
        this.id = id;
    }

    public double getXpos() {
        return xpos;
    }

    public void setXpos(float xpos) {
        this.xpos = xpos;
    }

    public double getYpos() {
        return ypos;
    }

    public void setYpos(float ypos) {
        this.ypos = ypos - height;
    }

    private String getDebugInfo() {
        return "\n[" + Double.toString(xpos) + "," + Double.toString(ypos) + "]\n" +
                "[" + Double.toString(width) + "," + Double.toString(height) + "]\n" +
                "[" + Double.toString(calcLocY()) + "]";
    }
    private double calcLocY() {
        //return 1.0 - height - height -height;
        return 0.5;
    }
    @Override
    public Element getXmlElement(Document document) {
        Element root = document.createElement("Shape");
        root.setAttribute("ID", Integer.toString(id));
        root.setAttribute("NameU", name);
        root.setAttribute("Name", name);
        root.setAttribute("Type", "Shape");
        root.setAttribute("Master", "2");

        XmlTool xmltool = new XmlTool(document);
        root.appendChild(xmltool.getElement("Cell", "N", "PinX", "V", Double.toString(xpos)));
        root.appendChild(xmltool.getElement("Cell", "N", "PinY", "V", Double.toString(ypos)));

        Element section = xmltool.getElement("Section", "N", "Character");
        Element rowlang = xmltool.getElement("Row", "IX", "0");
        rowlang.appendChild(xmltool.getElement("Cell", "N", "LangID", "V", "en-US"));
        section.appendChild(rowlang);
        root.appendChild(section);

        Element textTitle = xmltool.getElement("Text");
        textTitle.appendChild(xmltool.getElement("cp", "IX", "0"));
        textTitle.setTextContent(getTitle());
        root.appendChild(textTitle);

        root.appendChild(xmltool.getElement("Cell", "N", "Width", "V", Double.toString(width)));
        root.appendChild(xmltool.getElement("Cell", "N", "Height", "V", Double.toString(height)));
        root.appendChild(xmltool.getElement("Cell", "N", "LocPinY", "V", Double.toString(calcLocY()), "F", "Inh"));

        // connection section
        Element connectSection = xmltool.getElement("Section", "N", "Connection");
        Element cr0 = xmltool.getElement("Row", "IX", "0");
        cr0.appendChild(xmltool.getElement("Cell", "N", "Y", "V", "0", "F", "Inh"));
        connectSection.appendChild(cr0);

        Element cr1 = xmltool.getElement("Row", "IX", "1");
        cr1.appendChild(xmltool.getElement("Cell", "N", "Y", "V", Double.toString(calcLocY()), "F", "Inh"));
        connectSection.appendChild(cr1);

        Element cr2 = xmltool.getElement("Row", "IX", "2");
        cr2.appendChild(xmltool.getElement("Cell", "N", "Y", "V", "1", "F", "Inh"));
        connectSection.appendChild(cr2);

        Element cr3 = xmltool.getElement("Row", "IX", "3");
        cr3.appendChild(xmltool.getElement("Cell", "N", "Y", "V", Double.toString(calcLocY()), "F", "Inh"));
        connectSection.appendChild(cr3);

        Element cr4 = xmltool.getElement("Row", "IX", "4");
        cr4.appendChild(xmltool.getElement("Cell", "N", "Y", "V", Double.toString(calcLocY()), "F", "Inh"));
        connectSection.appendChild(cr4);

        root.appendChild(connectSection);
        //~connection section

        // geometry
        Element geometry = xmltool.getElement("Section", "N", "Geometry", "IX", "0");

        Element gr2 = xmltool.getElement("Row", "T", "LineTo", "IX", "2");
        gr2.appendChild(xmltool.getElement("Cell", "N", "X", "V", Double.toString(width), "F", "Inh"));
        geometry.appendChild(gr2);

        Element gr3 = xmltool.getElement("Row", "T", "LineTo", "IX", "3");
        gr3.appendChild(xmltool.getElement("Cell", "N", "X", "V", Double.toString(width), "F", "Inh"));
        gr3.appendChild(xmltool.getElement("Cell", "N", "Y", "V", Double.toString(height), "F", "Inh"));
        geometry.appendChild(gr3);

        Element gr4 = xmltool.getElement("Row", "T", "LineTo", "IX", "4");
        gr4.appendChild(xmltool.getElement("Cell", "N", "Y", "V", Double.toString(height), "F", "Inh"));
        geometry.appendChild(gr4);

        root.appendChild(geometry);
/*
        Element elPosX = document.createElement("Cell");
        elPosX.setAttribute("N", "PinX");
        elPosX.setAttribute("V", Double.toString(xpos));
        root.appendChild(elPosX);

        Element elPosY = document.createElement("Cell");
        elPosY.setAttribute("N", "PinY");
        elPosY.setAttribute("V", Double.toString(ypos));
        root.appendChild(elPosY);

        if (title != null && !title.isEmpty()) {
            Element section = document.createElement("Section");
            section.setAttribute("N", "Character");
            root.appendChild(section);
            Element row = document.createElement("Row");
            row.setAttribute("IX", "0");
            section.appendChild(row);
            Element cell = document.createElement("Cell");
            cell.setAttribute("N", "LangID");
            cell.setAttribute("V", "en-US");
            row.appendChild(cell);

            Element textTitle = document.createElement("Text");
            textTitle.setTextContent(getTitle());
            Element textcp = document.createElement("cp");
            textcp.setAttribute("IX", "0");
            textTitle.appendChild(textcp);
            root.appendChild(textTitle);
        }

        if (width > 0.0 || height > 0.0) {
            if (width > 0.0) {
                Element cellWidth = document.createElement("Cell");
                cellWidth.setAttribute("N", "Width");
                cellWidth.setAttribute("V", Double.toString(width));
                root.appendChild(cellWidth);
            }
            if (height > 0.0) {
                Element cellHeight = document.createElement("Cell");
                cellHeight.setAttribute("N", "Height");
                cellHeight.setAttribute("V", Double.toString(height));
                root.appendChild(cellHeight);
            }

//            Element cellLocPinX = document.createElement("Cell");
//            cellLocPinX.setAttribute("N", "LocPinX");
//            cellLocPinX.setAttribute("V", "0.4");
//            cellLocPinX.setAttribute("F", "Inh");
//            root.appendChild(cellLocPinX);

            if (height > 0.0) {
                Element cellLocPinY = document.createElement("Cell");
                cellLocPinY.setAttribute("N", "LocPinY");
                cellLocPinY.setAttribute("V", Double.toString(calcLocY()));
                cellLocPinY.setAttribute("F", "Inh");
                root.appendChild(cellLocPinY);
            }

            Element conn = document.createElement("Section");
            conn.setAttribute("N", "Connection");
            Element rc1 = document.createElement("Row");
            rc1.setAttribute("IX", "0");
            conn.appendChild(rc1);
            Element rcc1 = document.createElement("Cell");
            rcc1.setAttribute("N", "Y");
            rcc1.setAttribute("V", "0");
            rcc1.setAttribute("F", "Inh");
            rc1.appendChild(rcc1);

            Element rc2 = document.createElement("Row");
            rc2.setAttribute("IX", "1");
            conn.appendChild(rc2);
            Element rcc2 = document.createElement("Cell");
            rcc2.setAttribute("N", "Y");
            rcc2.setAttribute("V", Double.toString(calcLocY()));
            rcc2.setAttribute("F", "Inh");
            rc2.appendChild(rcc2);

            Element rc3 = document.createElement("Row");
            rc3.setAttribute("IX", "2");
            conn.appendChild(rc3);
            Element rcc3 = document.createElement("Cell");
            rcc3.setAttribute("N", "Y");
            rcc3.setAttribute("V", "1");
            rcc3.setAttribute("F", "Inh");
            rc3.appendChild(rcc3);

            Element rc4 = document.createElement("Row");
            rc4.setAttribute("IX", "3");
            conn.appendChild(rc4);
            Element rcc4 = document.createElement("Cell");
            rcc4.setAttribute("N", "Y");
            rcc4.setAttribute("V", Double.toString(calcLocY()));
            rcc4.setAttribute("F", "Inh");
            rc4.appendChild(rcc4);

            Element rc5 = document.createElement("Row");
            rc5.setAttribute("IX", "4");
            conn.appendChild(rc5);

            Element rcc5 = document.createElement("Cell");
            rcc5.setAttribute("N", "Y");
            rcc5.setAttribute("V", Double.toString(calcLocY()));
            rcc5.setAttribute("F", "Inh");
            rc5.appendChild(rcc5);

            root.appendChild(conn);


            Element geometry = document.createElement("Section");
            geometry.setAttribute("N", "Geometry");
            geometry.setAttribute("IX", "0");
            root.appendChild(geometry);

            if (width > 0.0) {
                Element row1 = document.createElement("Row");
                row1.setAttribute("T", "LineTo");
                row1.setAttribute("IX", "2");
                geometry.appendChild(row1);

                Element cell1 = document.createElement("Cell");
                cell1.setAttribute("N", "X");
                cell1.setAttribute("V", Double.toString(width));
                cell1.setAttribute("F", "Inh");
                row1.appendChild(cell1);
            }

            Element row2 = document.createElement("Row");
            row2.setAttribute("T", "LineTo");
            row2.setAttribute("IX", "3");
            geometry.appendChild(row2);

            if (width > 0.0) {
                Element cell2 = document.createElement("Cell");
                cell2.setAttribute("N", "X");
                cell2.setAttribute("V", Double.toString(width));
                cell2.setAttribute("F", "Inh");
                row2.appendChild(cell2);
            }
            if (height > 0.0) {
                Element cell3 = document.createElement("Cell");
                cell3.setAttribute("N", "Y");
                cell3.setAttribute("V", Double.toString(height));
                cell3.setAttribute("F", "Inh");
                row2.appendChild(cell3);
            }

            if (height > 0.0) {
                Element row3 = document.createElement("Row");
                row3.setAttribute("T", "LineTo");
                row3.setAttribute("IX", "4");
                geometry.appendChild(row3);

                Element cell4 = document.createElement("Cell");
                cell4.setAttribute("N", "Y");
                cell4.setAttribute("V", Double.toString(height));
                cell4.setAttribute("F", "Inh");
                row3.appendChild(cell4);
            }
        }*/

        return root;
    }
}
