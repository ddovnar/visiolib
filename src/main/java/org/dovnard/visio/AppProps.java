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
 * Visio project
 * Author: Dmitriy Dovnar
 * Date: 07.11.19
 * Time: 1:34
 */
public class AppProps implements XmlExportable {
    private final String docSchema = "http://schemas.openxmlformats.org/officeDocument/2006/extended-properties";
    private final String docSchemaVt = "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes";

    private String company;
    private String manager;
    private boolean linksUpToDate;
    private boolean sharedDoc;
    private boolean hyperlinksChanged;

    public AppProps() {
        manager = "";
        company = "";
        linksUpToDate = false;
        sharedDoc = false;
        hyperlinksChanged = false;
    }
    public boolean isHyperlinksChanged() {
        return hyperlinksChanged;
    }

    public void setHyperlinksChanged(boolean hyperlinksChanged) {
        this.hyperlinksChanged = hyperlinksChanged;
    }

    public String getCompany() {
        return company;
    }

    public void setCompany(String company) {
        this.company = company;
    }

    public String getManager() {
        return manager;
    }

    public void setManager(String manager) {
        this.manager = manager;
    }

    public boolean isLinksUpToDate() {
        return linksUpToDate;
    }

    public void setLinksUpToDate(boolean linksUpToDate) {
        this.linksUpToDate = linksUpToDate;
    }

    public boolean isSharedDoc() {
        return sharedDoc;
    }

    public void setSharedDoc(boolean sharedDoc) {
        this.sharedDoc = sharedDoc;
    }

    @Override
    public Document toXml() throws ParserConfigurationException {
        DocumentBuilderFactory docFactory = DocumentBuilderFactory.newInstance();

        DocumentBuilder docBuilder = docFactory.newDocumentBuilder();
        Document document = docBuilder.newDocument();
        document.setXmlStandalone(true);

        Element root = document.createElement("Properties");
        root.setAttribute("xmlns", docSchema);
        root.setAttribute("xmlns:vt", docSchemaVt);
        document.appendChild(root);

        Element template = document.createElement("Template");
        root.appendChild(template);

        Element application = document.createElement("Application");
        application.setTextContent("Microsoft Visio");
        root.appendChild(application);

        Element scaleCrop = document.createElement("ScaleCrop");
        scaleCrop.setTextContent("false");
        root.appendChild(scaleCrop);

        Element headingPairs = document.createElement("HeadingPairs");
            Element vtVector1 = document.createElement("vt:vector");
            vtVector1.setAttribute("size", "2");
            vtVector1.setAttribute("baseType", "variant");
                Element vtVariant1 = document.createElement("vt:variant");
                    Element item1 = document.createElement("vt:lpstr");
                    item1.setTextContent("Pages");
                    vtVariant1.appendChild(item1);
                vtVector1.appendChild(vtVariant1);

                Element vtVariant2 = document.createElement("vt:variant");
                    Element item2 = document.createElement("vt:i4");
                    item2.setTextContent("2");
                    vtVariant2.appendChild(item2);
                vtVector1.appendChild(vtVariant2);
            headingPairs.appendChild(vtVector1);
        root.appendChild(headingPairs);

        Element titlesOfParts = document.createElement("TitlesOfParts");
            Element vtVector2 = document.createElement("vt:vector");
            vtVector2.setAttribute("size", "2");
            vtVector2.setAttribute("baseType", "lpstr");
                Element item3 = document.createElement("vt:lpstr");
                item3.setTextContent("MyFirstPage");
                vtVector2.appendChild(item3);

                Element item4 = document.createElement("vt:lpstr");
                item4.setTextContent("MySecondPage");
                vtVector2.appendChild(item4);
            titlesOfParts.appendChild(vtVector2);
        root.appendChild(titlesOfParts);

        Element manager = document.createElement("Manager");
        manager.setTextContent(this.manager);
        root.appendChild(manager);

        Element company = document.createElement("Company");
        company.setTextContent(this.company);
        root.appendChild(company);

        Element linksUpToDate = document.createElement("LinksUpToDate");
        linksUpToDate.setTextContent(Boolean.toString(this.linksUpToDate));
        //linksUpToDate.setTextContent("false");
        root.appendChild(linksUpToDate);

        Element sharedDoc = document.createElement("SharedDoc");
        sharedDoc.setTextContent(Boolean.toString(this.sharedDoc));
        //sharedDoc.setTextContent("false");
        root.appendChild(sharedDoc);

        Element hyperLinkBase = document.createElement("HyperlinkBase");
        root.appendChild(hyperLinkBase);

        Element hyperlinksChanged = document.createElement("HyperlinksChanged");
//        hyperlinksChanged.setTextContent("false");
        hyperlinksChanged.setTextContent(Boolean.toString(this.hyperlinksChanged));
        root.appendChild(hyperlinksChanged);

        Element appVersion = document.createElement("AppVersion");
        appVersion.setTextContent("15.0000");
        root.appendChild(appVersion);

        return document;
    }

    @Override
    public String getFileName() {
        return "/docProps/" + "app.xml";
    }

    @Override
    public void writeToFile(String path) throws IOException, TransformerException, ParserConfigurationException {
        XmlFileWriter.writeToFile(path, getFileName(), toXml());
//        String fileName = getFileName();
//
//        if (!fileName.endsWith(".xml")) {
//            fileName += ".xml";
//        }
//
//        if (path != null) {
//            if (!path.endsWith("/") && !fileName.startsWith("/")) {
//                path += "/";
//            }
//        }
//        fileName = path + fileName;
//
//        DOMSource source = new DOMSource(toXml());
//        TransformerFactory transformerFactory = TransformerFactory.newInstance();
//        Transformer transformer = transformerFactory.newTransformer();
//        transformer.setOutputProperty(OutputKeys.STANDALONE, "yes");
//
//        StreamResult result = new StreamResult(new File(fileName));
//        transformer.transform(source, result);
    }

    @Override
    public void writeToFile(String path, String filename) throws IOException, TransformerException, ParserConfigurationException {
        XmlFileWriter.writeToFile(path, filename, toXml());
    }

}
