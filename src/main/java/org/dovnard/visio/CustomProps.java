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

/**
 * User: DovnarDmitriy
 * Date: 07.11.19
 * Time: 15:31
 */
public class CustomProps implements XmlExportable {
    @Override
    public Document toXml() throws ParserConfigurationException {
        DocumentBuilderFactory docFactory = DocumentBuilderFactory.newInstance();

        DocumentBuilder docBuilder = docFactory.newDocumentBuilder();
        Document document = docBuilder.newDocument();
        document.setXmlStandalone(true);

        Element root = document.createElement("Properties");
        root.setAttribute("xmlns", "http://schemas.openxmlformats.org/officeDocument/2006/custom-properties");
        root.setAttribute("xmlns:vt", "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes");
        document.appendChild(root);

        Element prop1 = document.createElement("property");
        prop1.setAttribute("fmtid", "{D5CDD505-2E9C-101B-9397-08002B2CF9AE}");
        prop1.setAttribute("pid", "2");
        prop1.setAttribute("name", "_VPID_ALTERNATENAMES");
        root.appendChild(prop1);
            Element val1 = document.createElement("vt:lpwstr");
            prop1.appendChild(val1);

        Element prop2 = document.createElement("property");
        prop2.setAttribute("fmtid", "{D5CDD505-2E9C-101B-9397-08002B2CF9AE}");
        prop2.setAttribute("pid", "3");
        prop2.setAttribute("name", "BuildNumberCreated");
        root.appendChild(prop2);
            Element val2 = document.createElement("vt:i4");
            val2.setTextContent("1006637809");
            prop2.appendChild(val2);

        Element prop3 = document.createElement("property");
        prop3.setAttribute("fmtid", "{D5CDD505-2E9C-101B-9397-08002B2CF9AE}");
        prop3.setAttribute("pid", "4");
        prop3.setAttribute("name", "BuildNumberEdited");
        root.appendChild(prop3);
            Element val3 = document.createElement("vt:i4");
            val3.setTextContent("1006637809");
            prop3.appendChild(val3);

        Element prop4 = document.createElement("property");
        prop4.setAttribute("fmtid", "{D5CDD505-2E9C-101B-9397-08002B2CF9AE}");
        prop4.setAttribute("pid", "5");
        prop4.setAttribute("name", "IsMetric");
        root.appendChild(prop4);
            Element val4 = document.createElement("vt:boot");
            val4.setTextContent("true");
            prop4.appendChild(val4);

        Element prop5 = document.createElement("property");
        prop5.setAttribute("fmtid", "{D5CDD505-2E9C-101B-9397-08002B2CF9AE}");
        prop5.setAttribute("pid", "4");
        prop5.setAttribute("name", "TimeEdited");
        root.appendChild(prop5);
            Element val5 = document.createElement("vt:filetime");
            val5.setTextContent("2016-08-31T19:23:09Z");
            prop5.appendChild(val5);

        return document;
    }

    @Override
    public String getFileName() {
        return "/docProps/" + "custom.xml";
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
