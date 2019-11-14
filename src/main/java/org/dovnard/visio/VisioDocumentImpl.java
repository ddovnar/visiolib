package org.dovnard.visio;

import org.w3c.dom.Document;

import javax.xml.parsers.ParserConfigurationException;
import javax.xml.transform.*;
import javax.xml.transform.dom.DOMSource;
import javax.xml.transform.stream.StreamResult;
import java.io.File;
import java.io.IOException;
import java.nio.file.Files;
import java.util.ArrayList;
import java.util.List;

/**
 * Visio project
 * Author: Dmitriy Dovnar
 * Date: 06.11.19
 * Time: 23:00
 */
public class VisioDocumentImpl implements VisioDocument {
    private Relationship packageRels;
    private Relationship documentRels;
    private PageCollection pages;

    private AppProps appProps;
    private CoreProps coreProps;
    private CustomProps customProps;

    private WindowsProps windowsProps;
    private DocumentProps documentProps;

    private ContentTypes contentTypes;

    public VisioDocumentImpl() throws Exception {
        pages = new PageCollection();
        packageRels = new Relationship();
        documentRels = new Relationship();

        appProps = new AppProps();
        coreProps = new CoreProps();
        customProps = new CustomProps();

        windowsProps = new WindowsProps();
        documentProps = new DocumentProps();

        contentTypes = new ContentTypes();

        addPackageRelations();
        addDocumentRalations();
    }

    private void addPackageRelations() throws Exception {
        packageRels.addRelation("rId3", "docProps/core.xml", "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties");
        packageRels.addRelation("rId2", "docProps/thumbnail.emf", "http://schemas.openxmlformats.org/package/2006/relationships/metadata/thumbnail");
        packageRels.addRelation("rId1", "visio/document.xml", "http://schemas.microsoft.com/visio/2010/relationships/document");
        packageRels.addRelation("rId5", "docProps/custom.xml", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/custom-properties");
        packageRels.addRelation("rId4", "docProps/app.xml", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties");
    }

    private void addDocumentRalations() throws Exception {
        documentRels.addRelation("rId2", "windows.xml", "http://schemas.microsoft.com/visio/2010/relationships/windows");
        documentRels.addRelation("rId1", "pages/pages.xml", "http://schemas.microsoft.com/visio/2010/relationships/pages");
        documentRels.addRelation("rId3", "masters/masters.xml", "http://schemas.microsoft.com/visio/2010/relationships/masters");
    }

    public void test() {
        System.out.println("Test");
    }

    public VisioPage createPage(String pageName) throws Exception {
        VisioPage newPage = pages.createPage(pageName, contentTypes);
        return newPage;
    }

    public List<VisioPage> getPageList() {
        return pages.getPageList();
    }

    private void copyStaticFiles(String targetPath) throws IOException {
        File shape1 = new File(getClass().getClassLoader().getResource("./shapes/thumbnail.emf").getFile());
        File targetFile = new File(targetPath + "/docProps/thumbnail.emf");
        if (shape1.exists() && !targetFile.exists()) {
            Files.copy(shape1.toPath(), targetFile.toPath());
        }

        File masters = new File(getClass().getClassLoader().getResource("./masters/masters.xml").getFile());
        File targetMasters = new File(targetPath + "/visio/masters/masters.xml");
        if (masters.exists() && !targetMasters.exists()) {
            File folder = new File(targetPath + "/visio/masters/");
            if (!folder.exists()) {
                folder.mkdir();
            }
            Files.copy(masters.toPath(), targetMasters.toPath());
        }

        File master1 = new File(getClass().getClassLoader().getResource("./masters/master1.xml").getFile());
        File targetMaster1 = new File(targetPath + "/visio/masters/master1.xml");
        if (master1.exists() && !targetMaster1.exists()) {
            Files.copy(master1.toPath(), targetMaster1.toPath());
        }

        File master2 = new File(getClass().getClassLoader().getResource("./masters/master2.xml").getFile());
        File targetMaster2 = new File(targetPath + "/visio/masters/master2.xml");
        if (master2.exists() && !targetMaster2.exists()) {
            Files.copy(master2.toPath(), targetMaster2.toPath());
        }

        File mastersR = new File(getClass().getClassLoader().getResource("./masters/_rels/masters.xml.rels").getFile());
        System.out.println("DDD: " + mastersR.getAbsolutePath());
        File targetMastersR = new File(targetPath + "/visio/masters/_rels/masters.xml.rels");
        if (mastersR.exists() && !targetMastersR.exists()) {
            System.out.println("Copy rels master");
            File folderR = new File(targetPath + "/visio/masters/_rels/");
            if (!folderR.exists()) {
                folderR.mkdir();
            }
            Files.copy(mastersR.toPath(), targetMastersR.toPath());
        }

        File mastersPR = new File(getClass().getClassLoader().getResource("./page_rels/page1.xml.rels").getFile());
        File targetMastersPR = new File(targetPath + "/visio/pages/_rels/page1.xml.rels");
        if (mastersPR.exists() && !targetMastersPR.exists()) {
            System.out.println("Copy rels pages");
            File folderPR = new File(targetPath + "/visio/pages/_rels/");
            if (!folderPR.exists()) {
                folderPR.mkdir();
            }
            Files.copy(mastersPR.toPath(), targetMastersPR.toPath());
        }
    }

    public void createVisioFile(String fileName) throws Exception {
        String tmpFolder = "tmp";

        packageRels.writeToFile(tmpFolder);
        documentRels.writeToFile(tmpFolder + "/visio/_rels", "document.xml.rels");
        appProps.writeToFile(tmpFolder);
        coreProps.writeToFile(tmpFolder);
        customProps.writeToFile(tmpFolder);

        pages.writeToFile(tmpFolder);

        windowsProps.writeToFile(tmpFolder);
        documentProps.writeToFile(tmpFolder);

        contentTypes.writeToFile(tmpFolder);

        copyStaticFiles(tmpFolder);

        ZipCreator zip = new ZipCreator();
        zip.createZip("c:\\work\\test3\\study\\visiolib\\tmp", "output/" + fileName);
    }

    @Override
    public void writeToFile(String fileName) throws Exception {
        String tmpFolder = "tmp";

        packageRels.writeToFile(tmpFolder);
        documentRels.writeToFile(tmpFolder + "/visio/_rels", "document.xml.rels");
        appProps.writeToFile(tmpFolder);
        coreProps.writeToFile(tmpFolder);
        customProps.writeToFile(tmpFolder);

        copyStaticFiles(tmpFolder);

        pages.writeToFile(tmpFolder);

        windowsProps.writeToFile(tmpFolder);
        documentProps.writeToFile(tmpFolder);

        contentTypes.writeToFile(tmpFolder);

//        System.out.println("");
//        writeToConsole(packageRels.toXml());
//
//        writeXmlToFile(tmpFolder + "/" + fileName, packageRels.toXml());
//
//        System.out.println("");
//        writeToConsole(documentRels.toXml());
//
//        System.out.println("");
//        writeToConsole(appProps.toXml());
//
//        System.out.println("");
//        writeToConsole(coreProps.toXml());
//
//        System.out.println("");
//        writeToConsole(customProps.toXml());
//
//        System.out.println("");
//        writeToConsole(pages.toXml());
    }

    private void writeXmlToFile(String fileName, Document xmlDocument) throws TransformerException {
        if (!fileName.endsWith(".xml")) {
            fileName += ".xml";
        }
        DOMSource source = new DOMSource(xmlDocument);
        TransformerFactory transformerFactory = TransformerFactory.newInstance();
        Transformer transformer = transformerFactory.newTransformer();
        transformer.setOutputProperty(OutputKeys.STANDALONE, "yes");

        StreamResult result = new StreamResult(new File(fileName));
        transformer.transform(source, result);
    }
    @Override
    public void setManager(String manager) {
        appProps.setManager(manager);
    }

    @Override
    public void setCompany(String company) {
        appProps.setCompany(company);
    }

    @Override
    public void setLanguage(String language) {
        coreProps.setLanguage(language);
    }

    public void writeToConsole(Document document) throws TransformerException {
        DOMSource source = new DOMSource(document);
        StreamResult result = new StreamResult(System.out);

        TransformerFactory transformerFactory = TransformerFactory.newInstance();
        Transformer transformer = transformerFactory.newTransformer();
        transformer.setOutputProperty(OutputKeys.STANDALONE, "yes");
        transformer.transform(source, result);
    }
}
