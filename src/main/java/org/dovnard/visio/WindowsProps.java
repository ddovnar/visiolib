package org.dovnard.visio;

import org.w3c.dom.Document;
import org.w3c.dom.Element;

import javax.xml.parsers.DocumentBuilder;
import javax.xml.parsers.DocumentBuilderFactory;
import javax.xml.parsers.ParserConfigurationException;
import javax.xml.transform.TransformerException;
import java.io.IOException;

/**
 * User: DovnarDmitriy
 * Date: 08.11.19
 * Time: 14:03
 */
public class WindowsProps implements XmlExportable {
    @Override
    public Document toXml() throws ParserConfigurationException {
        DocumentBuilderFactory docFactory = DocumentBuilderFactory.newInstance();

        DocumentBuilder docBuilder = docFactory.newDocumentBuilder();
        Document document = docBuilder.newDocument();
        document.setXmlStandalone(true);

        Element root = document.createElement("Windows");
        root.setAttribute("ClientWidth", "1366");
        root.setAttribute("ClientHeight", "563");
        root.setAttribute("xmlns", "http://schemas.microsoft.com/office/visio/2012/main");
        root.setAttribute("xmlns:r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
        root.setAttribute("xml:space", "preserve");

        Element window = document.createElement("Window");
        window.setAttribute("ID", "0");
        window.setAttribute("WindowType", "Drawing");
        window.setAttribute("WindowState", "1073741824");
        window.setAttribute("WindowLeft", "-8");
        window.setAttribute("WindowTop", "-30");
        window.setAttribute("WindowWidth", "1382");
        window.setAttribute("WindowHeight", "601");
        window.setAttribute("ContainerType", "Page");
        window.setAttribute("Page", "0");
        window.setAttribute("ViewScale", "0.82");
        window.setAttribute("ViewCenterX", "4.1275082550165");
        window.setAttribute("ViewCenterY", "8.5852171704343");
        root.appendChild(window);

        Element ShowRulers = document.createElement("ShowRulers");
        ShowRulers.setTextContent("1");
        window.appendChild(ShowRulers);

        Element ShowGrid = document.createElement("ShowGrid");
        ShowGrid.setTextContent("1");
        window.appendChild(ShowGrid);

        Element ShowPageBreaks = document.createElement("ShowPageBreaks");
        ShowPageBreaks.setTextContent("0");
        window.appendChild(ShowPageBreaks);

        Element ShowGuides = document.createElement("ShowGuides");
        ShowGuides.setTextContent("1");
        window.appendChild(ShowGuides);

        Element ShowConnectionPoints = document.createElement("ShowConnectionPoints");
        ShowConnectionPoints.setTextContent("1");
        window.appendChild(ShowConnectionPoints);

        Element GlueSettings = document.createElement("GlueSettings");
        GlueSettings.setTextContent("9");
        window.appendChild(GlueSettings);

        Element SnapSettings = document.createElement("SnapSettings");
        SnapSettings.setTextContent("65847");
        window.appendChild(SnapSettings);

        Element SnapExtensions = document.createElement("SnapExtensions");
        SnapExtensions.setTextContent("34");
        window.appendChild(SnapExtensions);

        window.appendChild(document.createElement("SnapAngles"));

        Element DynamicGridEnabled = document.createElement("DynamicGridEnabled");
        DynamicGridEnabled.setTextContent("1");
        window.appendChild(DynamicGridEnabled);

        Element TabSplitterPos = document.createElement("TabSplitterPos");
        TabSplitterPos.setTextContent("0.5");
        window.appendChild(TabSplitterPos);

        document.appendChild(root);

        return document;
    }

    @Override
    public String getFileName() {
        return "/visio/windows.xml";
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
