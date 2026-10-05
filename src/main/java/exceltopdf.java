import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.pdmodel.PDPage;
import org.apache.pdfbox.pdmodel.PDPageContentStream;
import org.apache.pdfbox.pdmodel.font.PDType1Font;
import org.apache.pdfbox.pdmodel.font.Standard14Fonts;
import org.apache.poi.ss.usermodel.*;

import java.io.FileInputStream;
import java.io.File;

public class exceltopdf {
    public static void main(String[] args) {
        String excelPath = "marksheet.xlsx";
        String pdfPath = "D:\\java\\marksheet.pdf";

        try {
            FileInputStream file = new FileInputStream(excelPath);
            Workbook workbook = WorkbookFactory.create(file);
            Sheet sheet = workbook.getSheetAt(0);

            PDDocument pdfDoc = new PDDocument();
            PDPage page = new PDPage();
            pdfDoc.addPage(page);

            PDPageContentStream contentStream = new PDPageContentStream(pdfDoc, page);
            DataFormatter formatter = new DataFormatter();

            contentStream.setFont(new PDType1Font(Standard14Fonts.FontName.HELVETICA_BOLD), 12);
            contentStream.beginText();
            contentStream.newLineAtOffset(50, 700);

            int yPosition = 700;

            for (Row row : sheet) {
                StringBuilder rowText = new StringBuilder();
                for (Cell cell : row) {
                    rowText.append(formatter.formatCellValue(cell)).append("   |   ");
                }

                contentStream.showText(rowText.toString());
                yPosition -= 20;
                contentStream.newLineAtOffset(0, -20);
            }

            contentStream.endText();
            contentStream.close();

            pdfDoc.save(pdfPath);
            pdfDoc.close();
            workbook.close();
            file.close();

            System.out.println("FILE HAS BEEN SUCCESSFULLY CONVERTED INTO PDF");
            System.out.println("THE PDF HAS DOWNLOADED IN THE Java FOLLLDER");
        } catch (Exception e) {
            e.printStackTrace();
        }
    }
}