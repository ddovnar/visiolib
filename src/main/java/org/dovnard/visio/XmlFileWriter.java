package org.dovnard.visio;

import org.w3c.dom.Document;

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
 * Date: 08.11.19
 * Time: 13:47
 */
public class XmlFileWriter {
    public static void writeToFile(String path, String filename, Document xmlDocument) throws IOException, TransformerException {
        String fileName = filename;
//        if (!fileName.endsWith(".xml")) {
//            fileName += ".xml";
//        }

        if (path != null) {
            if (!path.endsWith("/") && !fileName.startsWith("/")) {
                path += "/";
            }
        }
        fileName = path + fileName;

        DOMSource source = new DOMSource(xmlDocument);
        TransformerFactory transformerFactory = TransformerFactory.newInstance();
        Transformer transformer = transformerFactory.newTransformer();
        transformer.setOutputProperty(OutputKeys.STANDALONE, "yes");
        transformer.setOutputProperty(OutputKeys.ENCODING, "utf-8");

        File file = new File(fileName);
        file.getParentFile().mkdirs();
        StreamResult result = new StreamResult(file);
        transformer.transform(source, result);
    }
}
