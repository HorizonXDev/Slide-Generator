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
import java.util.Arrays;

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
                String title = sheet.getRow(i).getCell(0).getStringCellValue(); // Title column
                String content = sheet.getRow(i).getCell(1).getStringCellValue(); // Content column

                // Create a new slide from the template layout
                XSLFSlide slide = ppt.createSlide(ppt.getSlideMasters().get(0).getLayout(SlideLayout.TITLE_AND_CONTENT));
                System.out.println("Slide Created");
                // Set title placeholder text
                XSLFTextShape titlePlaceholder = slide.getPlaceholder(0);
                if (titlePlaceholder != null) {
                    titlePlaceholder.setText(title);
                }

                // Set content placeholder text
                XSLFTextShape contentPlaceholder = slide.getPlaceholder(1);
                if (contentPlaceholder != null) {
                    contentPlaceholder.setText(content);
                } else {
                    // If no content placeholder, create a text box
                    XSLFTextShape textBox = slide.createTextBox();
                    textBox.setAnchor(new java.awt.Rectangle(100, 100, 400, 300)); // Position
                    textBox.setText(content);
                    textBox.setTextAutofit(TextShape.TextAutofit.SHAPE);
                }

                String imagePath = sheet.getRow(i).getCell(2).getStringCellValue();
                if(imagePath != null && !imagePath.isEmpty()){
                    try(FileInputStream imageStream = new FileInputStream(imagePath)){
                        byte[] pictureData = imageStream.readAllBytes();
                        XSLFPictureData pictureIDX = ppt.addPicture(pictureData, PictureData.PictureType.PNG);
                        XSLFPictureShape picture = slide.createPicture(pictureIDX);
                        picture.setAnchor(new java.awt.Rectangle(50, 50, 400, 400));
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
