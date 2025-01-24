package io.horizon;
import org.apache.poi.sl.usermodel.PictureData;
import org.apache.poi.sl.usermodel.Placeholder;
import org.apache.poi.sl.usermodel.TextShape;
import org.apache.poi.xslf.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;


public class ExcelToPPT {
    public static void main(String[] args) {
        // Paths to input and output files
        String excelFilePath = "data.xlsx";
        String templateFilePath = "template.pptx";
        String outputFilePath = "output_presentation.pptx";

        try (FileInputStream excelFile = new FileInputStream(excelFilePath);
             FileInputStream templateFile = new FileInputStream(templateFilePath);
             XSSFWorkbook workbook = new XSSFWorkbook(excelFile);
             XMLSlideShow ppt = new XMLSlideShow(templateFile)) {

            // Read the first sheet of the Excel file
            XSSFSheet sheet = workbook.getSheetAt(0);

            // Iterate through the rows of the sheet
            for (int i = 1; i <= sheet.getLastRowNum(); i++) { // Skip the header row
                String name = sheet.getRow(i).getCell(0).getStringCellValue(); // Title column
                String genesis = sheet.getRow(i).getCell(1).getStringCellValue(); // Content column
                String knownFor = sheet.getRow(i).getCell(2).getStringCellValue();
                String nutsAbout = sheet.getRow(i).getCell(3).getStringCellValue();
                String sixImagePath = sheet.getRow(i).getCell(4).getStringCellValue();
                String twelveImagePath = sheet.getRow(i).getCell(5).getStringCellValue();

                // Create a new slide from the template layout
                XSLFSlide slide = ppt.createSlide(ppt.getSlideMasters().get(0).getLayout("Check"));
                System.out.println("Slide Created");
                // Set title placeholder text
                XSLFTextShape namePlaceholder = slide.getPlaceholder(0);
                if (namePlaceholder != null) {
                    namePlaceholder.setText(name);
                }

                // Set content placeholder text
                XSLFTextShape genesisPlaceholder = slide.getPlaceholder(1);
                if (genesisPlaceholder != null) {
                    genesisPlaceholder.setText(genesis);
                }

                XSLFTextShape knownForPlaceholder = slide.getPlaceholder(2);
                if (knownForPlaceholder != null) {
                    knownForPlaceholder.setText(knownFor);
                }
                XSLFTextShape nutsAboutPlaceholder = slide.getPlaceholder(3);
                if (nutsAboutPlaceholder != null) {
                    nutsAboutPlaceholder.setText(nutsAbout);
                }

                if(sixImagePath != null && !sixImagePath.isEmpty()){
                    try(FileInputStream imageStream = new FileInputStream(sixImagePath)){
                        byte[] pictureData = imageStream.readAllBytes();
                        XSLFPictureData pictureIDX = ppt.addPicture(pictureData, PictureData.PictureType.PNG);
                        XSLFPictureShape picture = slide.createPicture(pictureIDX);
                        picture.setAnchor(new java.awt.Rectangle(35, 35, 100, 100));
                    }
                }

                if(twelveImagePath != null && !twelveImagePath.isEmpty()){
                    try(FileInputStream imageStream = new FileInputStream(twelveImagePath)){
                        byte[] pictureData = imageStream.readAllBytes();
                        XSLFPictureData pictureIDX = ppt.addPicture(pictureData, PictureData.PictureType.PNG);
                        XSLFPictureShape picture = slide.createPicture(pictureIDX);
                        picture.setAnchor(new java.awt.Rectangle(585, 35, 100, 100));
                    }
                }
            }

            // Save the generated presentation
            try (FileOutputStream outputStream = new FileOutputStream(outputFilePath)) {
                ppt.write(outputStream);
            }

            System.out.println("Presentation created successfully!");

        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}
