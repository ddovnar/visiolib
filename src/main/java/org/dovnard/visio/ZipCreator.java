package org.dovnard.visio;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;
import java.util.zip.ZipEntry;
import java.util.zip.ZipOutputStream;

/**
 * User: DovnarDmitriy
 * Date: 11.11.19
 * Time: 11:17
 */
public class ZipCreator {
    private List<String> fileList;
    private String sourceFolder;

    public ZipCreator() {
        fileList = new ArrayList<>();
    }

    public void createZip(String sourceFolder, String outputFile) throws Exception {
        this.sourceFolder = sourceFolder;

        generateFileList(new File(sourceFolder));

        for( String f : fileList) {
            System.out.println("File: " + f);
        }

        zipIt(outputFile);
    }

    private String generateZipEntry(String file){
        return file.substring(sourceFolder.length()+1, file.length());
    }

    private void generateFileList(File node){

        //add file only
        if(node.isFile()){
            fileList.add(generateZipEntry(node.getAbsoluteFile().toString()));
        }

        if(node.isDirectory()){
            String[] subNote = node.list();
            for(String filename : subNote){
                generateFileList(new File(node, filename));
            }
        }
    }

    private void zipIt(String zipFile) throws Exception {

        byte[] buffer = new byte[1024];

        try{

            FileOutputStream fos = new FileOutputStream(zipFile);
            ZipOutputStream zos = new ZipOutputStream(fos);

            System.out.println("Output to Zip : " + zipFile);

            for(String file : this.fileList){

                System.out.println("File Added : " + file);
                ZipEntry ze= new ZipEntry(file);
                zos.putNextEntry(ze);

                FileInputStream in =
                        new FileInputStream(sourceFolder + File.separator + file);

                int len;
                while ((len = in.read(buffer)) > 0) {
                    zos.write(buffer, 0, len);
                }

                in.close();
            }

            zos.closeEntry();
            //remember close it
            zos.close();

            System.out.println("Done");
        }catch(IOException ex){
            ex.printStackTrace();
            throw new Exception(ex);
        }
    }
}
