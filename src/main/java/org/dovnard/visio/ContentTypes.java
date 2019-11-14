package org.dovnard.visio;

import org.w3c.dom.Document;
import org.w3c.dom.Element;

import javax.xml.parsers.DocumentBuilder;
import javax.xml.parsers.DocumentBuilderFactory;
import javax.xml.parsers.ParserConfigurationException;
import javax.xml.transform.TransformerException;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

/**
 * User: DovnarDmitriy
 * Date: 08.11.19
 * Time: 14:49
 */
public class ContentTypes implements XmlExportable {
    private List<String> registeredPages;

    public ContentTypes() {
        registeredPages = new ArrayList<>();
    }

    public void registerPage(String pageFileName) {
        registeredPages.add(pageFileName);
    }

    @Override
    public Document toXml() throws ParserConfigurationException {
        DocumentBuilderFactory docFactory = DocumentBuilderFactory.newInstance();

        DocumentBuilder docBuilder = docFactory.newDocumentBuilder();
        Document document = docBuilder.newDocument();
        document.setXmlStandalone(true);

        Element root = document.createElement("Types");
        root.setAttribute("xmlns", "http://schemas.openxmlformats.org/package/2006/content-types");
        document.appendChild(root);

        Element def1 = document.createElement("Default");
        def1.setAttribute("ContentType", "application/vnd.openxmlformats-package.relationships+xml");
        def1.setAttribute("Extension", "rels");
        root.appendChild(def1);

        Element def2 = document.createElement("Default");
        def2.setAttribute("ContentType", "application/xml");
        def2.setAttribute("Extension", "xml");
        root.appendChild(def2);

        Element def3 = document.createElement("Default");
        def3.setAttribute("ContentType", "image/x-emf");
        def3.setAttribute("Extension", "emf");
        root.appendChild(def3);

        Element over1 = document.createElement("Override");
        over1.setAttribute("ContentType", "application/vnd.openxmlformats-officedocument.custom-properties+xml");
        over1.setAttribute("PartName", "/docProps/custom.xml");
        root.appendChild(over1);

        Element over2 = document.createElement("Override");
        over2.setAttribute("ContentType", "application/vnd.ms-visio.pages+xml");
        over2.setAttribute("PartName", "/visio/pages/pages.xml");
        root.appendChild(over2);

        Element over3 = document.createElement("Override");
        over3.setAttribute("ContentType", "application/vnd.ms-visio.drawing.main+xml");
        over3.setAttribute("PartName", "/visio/document.xml");
        root.appendChild(over3);

        Element over4 = document.createElement("Override");
        over4.setAttribute("ContentType", "application/vnd.openxmlformats-officedocument.extended-properties+xml");
        over4.setAttribute("PartName", "/docProps/app.xml");
        root.appendChild(over4);

        Element over5 = document.createElement("Override");
        over5.setAttribute("ContentType", "application/vnd.ms-visio.windows+xml");
        over5.setAttribute("PartName", "/visio/windows.xml");
        root.appendChild(over5);

        Element over6 = document.createElement("Override");
        over6.setAttribute("ContentType", "application/vnd.openxmlformats-package.core-properties+xml");
        over6.setAttribute("PartName", "/docProps/core.xml");
        root.appendChild(over6);

        Element over7 = document.createElement("Override");
        over7.setAttribute("ContentType", "application/vnd.ms-visio.masters+xml");
        over7.setAttribute("PartName", "/visio/masters/masters.xml");
        root.appendChild(over7);

        Element over8 = document.createElement("Override");
        over8.setAttribute("ContentType", "application/vnd.ms-visio.master+xml");
        over8.setAttribute("PartName", "/visio/masters/master1.xml");
        root.appendChild(over8);

        Element over9 = document.createElement("Override");
        over9.setAttribute("ContentType", "application/vnd.ms-visio.master+xml");
        over9.setAttribute("PartName", "/visio/masters/master2.xml");
        root.appendChild(over9);

//        Element over7 = document.createElement("Override");
//        over7.setAttribute("ContentType", "application/vnd.ms-visio.page+xml");
//        over7.setAttribute("PartName", "/visio/pages/page1.xml");
//        root.appendChild(over7);
        Iterator<String> regpi = registeredPages.iterator();
        while (regpi.hasNext()) {
            Element overp = document.createElement("Override");
            overp.setAttribute("ContentType", "application/vnd.ms-visio.page+xml");
            overp.setAttribute("PartName", regpi.next());
            root.appendChild(overp);
        }

        return document;
    }

    @Override
    public String getFileName() {
        return "[Content_Types].xml";
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
