package org.dovnard.visio;

import org.w3c.dom.Document;
import org.w3c.dom.Element;

import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

/**
 * User: DovnarDmitriy
 * Date: 11.11.19
 * Time: 15:35
 */
public class ConnectorShape extends Shape {
    private final String v1 = "1";
    private final String v2 = "2";
    private final String v9 = "9";
    private final String v0_25 = "0.25";
    private final String vmin1 = "-1";
    private final String v0_125 = "0.125";
    private final String vmin0_5 = "-0.5";
    private final String v9_5 = "9.5";
    private final String v8_5 = "8.5";
    private double pinX;
    private double pinY;
    private double beginX;
    private double beginY;
    private double endX;
    private double endY;
    private double locPinX;
    private double locPinY;

    private List<DrawCommand> drawCmds;

    public ConnectorShape() {
        drawCmds = new ArrayList<>();
    }

    public void addDrawCmd(String cmd, double x, double y) {
        drawCmds.add(new DrawCommand(cmd, x, y));
    }
    public void setLocPinX(double v) {
        locPinX = v;
    }
    public void setLocPinY(double v) {
        locPinY = v;
    }
    public void setEndX(double v) {
        endX = v;
    }
    public void setEndY(double v) {
        endY = v;
    }
    public void setBeginX(double v) {
        beginX = v;
    }
    public void setBeginY(double v) {
        beginY = v;
    }
    public void setPinX(double v) {
        pinX = v;
    }
    public void setPinY(double v) {
        pinY = v;
    }
    private Element getPinX(Document document) {
        Element el = document.createElement("Cell");
        el.setAttribute("N", "PinX");
//        el.setAttribute("V", v1);
        el.setAttribute("V", Double.toString(pinX));
//        el.setAttribute("U", "MM");
        el.setAttribute("F", "Inh");
        return el;
    }
    private Element getPinY(Document document) {
        Element el = document.createElement("Cell");
        el.setAttribute("N", "PinY");
//        el.setAttribute("V", v9);
        el.setAttribute("V", Double.toString(pinY));
        el.setAttribute("F", "Inh");
        return el;
    }
    private Element getWidth(Document document) {
        Element el = document.createElement("Cell");
        el.setAttribute("N", "Width");
        el.setAttribute("V", v0_25);
        el.setAttribute("F", "GUARD(0.25DL)");
        return el;
    }
    private Element getHeight(Document document) {
        Element el = document.createElement("Cell");
        el.setAttribute("N", "Height");
        el.setAttribute("V", vmin1);
        el.setAttribute("F", "GUARD(EndY-BeginY)");
        return el;
    }
    private Element getLocPinX(Document document) {
        Element el = document.createElement("Cell");
        el.setAttribute("N", "LocPinX");
//        el.setAttribute("V", v0_125);
        el.setAttribute("V", Double.toString(locPinX));
        el.setAttribute("F", "Inh");
        return el;
    }
    private Element getLocPinY(Document document) {
        Element el = document.createElement("Cell");
        el.setAttribute("N", "LocPinY");
//        el.setAttribute("V", vmin0_5);
        el.setAttribute("V", Double.toString(locPinY));
        el.setAttribute("F", "Inh");
        return el;
    }
    private Element getBeginX(Document document) {
        Element el = document.createElement("Cell");
        el.setAttribute("N", "BeginX");
//        el.setAttribute("V", v1);
        el.setAttribute("V", Double.toString(beginX));
//        el.setAttribute("U", "MM");
        el.setAttribute("F", "PAR(PNT(Sheet.1!Connections.X1,Sheet.1!Connections.Y1))");
        return el;
    }
    private Element getBeginY(Document document) {
        Element el = document.createElement("Cell");
        el.setAttribute("N", "BeginY");
//        el.setAttribute("V", v9_5);
        el.setAttribute("V", Double.toString(beginY));
        el.setAttribute("F", "PAR(PNT(Sheet.1!Connections.X1,Sheet.1!Connections.Y1))");
        return el;
    }
    private Element getEndX(Document document) {
        Element el = document.createElement("Cell");
        el.setAttribute("N", "EndX");
//        el.setAttribute("V", v1);
        el.setAttribute("V", Double.toString(endX));
//        el.setAttribute("U", "MM");
        el.setAttribute("F", "PAR(PNT(Sheet.2!Connections.X3,Sheet.2!Connections.Y3))");
        return el;
    }
    private Element getEndY(Document document) {
        Element el = document.createElement("Cell");
        el.setAttribute("N", "EndY");
//        el.setAttribute("V", v8_5);
        el.setAttribute("V", Double.toString(endY));
        el.setAttribute("F", "PAR(PNT(Sheet.2!Connections.X3,Sheet.2!Connections.Y3))");
        return el;
    }
    private Element getLayerMember(Document document) {
        Element el = document.createElement("Cell");
        el.setAttribute("N", "LayerMember");
        el.setAttribute("V", "0");
        return el;
    }
    private Element getBeginTrigger(Document document) {
        Element el = document.createElement("Cell");
        el.setAttribute("N", "BeginTrigger");
        el.setAttribute("V", v2);
        el.setAttribute("F", "_XFTRIGGER(Sheet.1!EventXFMod)");
        return el;
    }
    private Element getEndTrigger(Document document) {
        Element el = document.createElement("Cell");
        el.setAttribute("N", "EndTrigger");
        el.setAttribute("V", v2);
        el.setAttribute("F", "_XFTRIGGER(Sheet.2!EventXFMod)");
        return el;
    }
    private Element getTxtPinX(Document document) {
        Element el = document.createElement("Cell");
        el.setAttribute("N", "TxtPinX");
        el.setAttribute("V", v0_125);
        el.setAttribute("F", "Inh");
        return el;
    }
    private Element getTxtPinY(Document document) {
        Element el = document.createElement("Cell");
        el.setAttribute("N", "TxtPinY");
        el.setAttribute("V", vmin0_5);
        el.setAttribute("F", "Inh");
        return el;
    }
    @Override
    public Element getXmlElement(Document document) {
        Element root = document.createElement("Shape");
        root.setAttribute("ID", Integer.toString(getId()));
        root.setAttribute("NameU", getName());
        root.setAttribute("Name", getName());
        root.setAttribute("Type", "Shape");
        root.setAttribute("Master", "6");

        root.appendChild(getPinX(document));
        root.appendChild(getPinY(document));
        root.appendChild(getWidth(document));
        root.appendChild(getHeight(document));
        root.appendChild(getLocPinX(document));
        root.appendChild(getLocPinY(document));
        root.appendChild(getBeginX(document));
        root.appendChild(getBeginY(document));
        root.appendChild(getEndX(document));
        root.appendChild(getEndY(document));
        root.appendChild(getLayerMember(document));
        root.appendChild(getBeginTrigger(document));
        root.appendChild(getEndTrigger(document));
        root.appendChild(getTxtPinX(document));
        root.appendChild(getTxtPinY(document));

        Element cofc = document.createElement("Cell");
        cofc.setAttribute("N", "ConFixedCode");
        cofc.setAttribute("V", "3");
        root.appendChild(cofc);

        root.appendChild(getGeometry(document));
        return root;
    }

    private Element getGeometry(Document document) {
        XmlTool xmltool = new XmlTool(document);

        Element el = document.createElement("Section");
        el.setAttribute("N", "Geometry");
        el.setAttribute("IX", "0");

        Iterator<DrawCommand> drawIter = drawCmds.iterator();
        int idx = 1;
        while (drawIter.hasNext()) {
            DrawCommand cmd = drawIter.next();
            Element r = xmltool.getElement("Row", "T", cmd.cmd, "IX", Integer.toString(idx));
            r.appendChild(xmltool.getElement("Cell", "N", "X", "V", Double.toString(cmd.x)));
            r.appendChild(xmltool.getElement("Cell", "N", "Y", "V", Double.toString(cmd.y)));
            el.appendChild(r);
            idx++;
        }

//        Element row1 = xmltool.getElement("Row", "T", "MoveTo", "IX", "1");
////        row1.appendChild(xmltool.getElement("Cell", "N", "X", "V", v0_125));
//        row1.appendChild(xmltool.getElement("Cell", "N", "X", "V", "0"));
//        row1.appendChild(xmltool.getElement("Cell", "N", "Y", "V", "0"));
//        el.appendChild(row1);
//
//        Element row2 = xmltool.getElement("Row", "T", "LineTo", "IX", "2");
////        row2.appendChild(xmltool.getElement("Cell", "N", "X", "V", v0_125));
////        row2.appendChild(xmltool.getElement("Cell", "N", "Y", "V", vmin1));
//        row2.appendChild(xmltool.getElement("Cell", "N", "X", "V", "0"));
//        row2.appendChild(xmltool.getElement("Cell", "N", "Y", "V", "-0.5"));
//        el.appendChild(row2);
//
//        Element row22 = xmltool.getElement("Row", "T", "LineTo", "IX", "3");
//        row22.appendChild(xmltool.getElement("Cell", "N", "X", "V", "-3.0"));
//        row22.appendChild(xmltool.getElement("Cell", "N", "Y", "V", "-0.5"));
//        el.appendChild(row22);
//
//        Element row222 = xmltool.getElement("Row", "T", "LineTo", "IX", "4");
//        row222.appendChild(xmltool.getElement("Cell", "N", "X", "V", "-3.0"));
//        row222.appendChild(xmltool.getElement("Cell", "N", "Y", "V", "-1.0"));
//        el.appendChild(row222);

//        Element row3 = xmltool.getElement("Row", "T", "LineTo", "IX", "4", "Del", "1");
//        el.appendChild(row3);

        return el;
    }

    class DrawCommand {
        private String cmd;
        private double x;
        private double y;
        public DrawCommand(String cmd, double x, double y) {
            this.cmd = cmd;
            this.x = x;
            this.y = y;
        }
    }
}
