package org.dovnard.visio;

import javax.xml.parsers.ParserConfigurationException;
import javax.xml.transform.TransformerException;
import java.io.IOException;
import java.time.LocalDateTime;
import java.util.List;

/**
 * Visio project
 * Author: Dmitriy Dovnar
 * Date: 07.11.19
 * Time: 0:13
 */
public interface VisioDocument {
    VisioPage createPage(String pageName) throws Exception;
    List<VisioPage> getPageList();
    void writeToFile(String fileName) throws Exception;
    void createVisioFile(String fileName) throws Exception;
    void setManager(String manager);
    void setCompany(String company);
    void setLanguage(String language);
}
