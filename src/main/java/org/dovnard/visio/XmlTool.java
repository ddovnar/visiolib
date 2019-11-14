package org.dovnard.visio;

import org.w3c.dom.Document;
import org.w3c.dom.Element;

/**
 * User: DovnarDmitriy
 * Date: 12.11.19
 * Time: 12:39
 */
public class XmlTool {
    private Document doc;

    public XmlTool(Document doc) {
        this.doc = doc;
    }
    public Element getElement(String tag) {
        Element el = doc.createElement(tag);
        return el;
    }
    public Element getElement(String tag, String p1name, String p1val) {
        Element el = doc.createElement(tag);
        el.setAttribute(p1name, p1val);
        return el;
    }
    public Element getElement(String tag, String p1name, String p1val, String p2name, String p2val) {
        Element el = doc.createElement(tag);
        el.setAttribute(p1name, p1val);
        el.setAttribute(p2name, p2val);
        return el;
    }
    public Element getElement(String tag, String p1name, String p1val, String p2name, String p2val, String p3name, String p3val) {
        Element el = doc.createElement(tag);
        el.setAttribute(p1name, p1val);
        el.setAttribute(p2name, p2val);
        el.setAttribute(p3name, p3val);
        return el;
    }
}
