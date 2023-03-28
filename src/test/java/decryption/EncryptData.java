package decryption;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Base64;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.testng.annotations.Test;

public class EncryptData {
@Test
    public  void main() throws IOException {
        // Open the input Excel file
        File inputFile = new File("C:\\Users\\DELL\\eclipse-workspace\\S.Grid\\Book.xlsx");
        FileInputStream fis = new FileInputStream(inputFile);

        // Create a Workbook object
        Workbook workbook = WorkbookFactory.create(fis);

        // Get the sheet with the data to encrypt
        Sheet sheet = workbook.getSheetAt(0);

        // Loop through all the rows in the sheet
        for (Row row : sheet) {
            // Get the cell in column 2 (index 1)
            Cell cellToEncrypt = row.getCell(1);

            // Check if the cell is not null and not empty
            if (cellToEncrypt != null && !cellToEncrypt.getStringCellValue().isEmpty()) {
                // Get the cell value
                String dataToEncrypt = cellToEncrypt.getStringCellValue();

                // Encode the data
                byte[] encodedBytes = Base64.getEncoder().encode(dataToEncrypt.getBytes());
                String encryptedData = new String(encodedBytes);
                System.out.print("the encrypt data" + encryptedData);

                // Write the encrypted data to the cell in column 3 (index 2) with column name "Encrypted Data"
                Cell encryptedCell = row.createCell(2);
                encryptedCell.setCellValue(encryptedData);
                sheet.getRow(0).createCell(2).setCellValue("Encrypted Data");
            }
        }

        // Save the updated Excel file
        FileOutputStream fos = new FileOutputStream(inputFile);
        workbook.write(fos);
        workbook.close();
        fis.close();
        fos.close();
    }
}
