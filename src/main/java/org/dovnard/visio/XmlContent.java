package org.dovnard.visio;

import org.w3c.dom.Document;
import org.w3c.dom.Element;

/**
 * User: DovnarDmitriy
 * Date: 08.11.19
 * Time: 16:48
 */
public interface XmlContent {
    Element getXmlElement(Document document);
}
