package org.dovnard.visio;

/**
 * Visio project
 * Author: Dmitriy Dovnar
 * Date: 07.11.19
 * Time: 0:21
 */
public class VisioDocumentBuilder {
    public static VisioDocument newDocument() throws Exception {
        return new VisioDocumentImpl();
    }
}
