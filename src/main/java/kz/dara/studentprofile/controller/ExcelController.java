package kz.dara.studentprofile.controller;

import kz.dara.studentprofile.util.AesEncryptDecrypt;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.http.HttpHeaders;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.*;
import org.springframework.web.multipart.MultipartFile;

import java.io.ByteArrayOutputStream;
import java.io.IOException;

@RestController
@RequestMapping("/api/excel")
public class ExcelController {

    @PostMapping("/decrypt")
    public ResponseEntity<byte[]> decryptFile(@RequestParam("file") MultipartFile file) throws IOException {
        Workbook workbook = new XSSFWorkbook(file.getInputStream());
        Sheet sheet = workbook.getSheetAt(0);

        for (Row row : sheet) {
            Cell iinCell = row.getCell(1);
            if (iinCell != null && iinCell.getCellType() == CellType.STRING) {
                String encryptedIin = iinCell.getStringCellValue();
                String decryptedIin = AesEncryptDecrypt.decrypt(encryptedIin);
                if (decryptedIin != null) {
                    iinCell.setCellValue(decryptedIin);
                }
            }
        }

        ByteArrayOutputStream out = new ByteArrayOutputStream();
        workbook.write(out);
        workbook.close();

        HttpHeaders headers = new HttpHeaders();
        headers.add(HttpHeaders.CONTENT_DISPOSITION, "attachment; filename=decrypted_iin.xlsx");

        return ResponseEntity.ok()
                .headers(headers)
                .contentType(MediaType.APPLICATION_OCTET_STREAM)
                .body(out.toByteArray());
    }
    @PostMapping("/encrypt")
    public ResponseEntity<byte[]> encryptFile(@RequestParam("file") MultipartFile file) throws IOException {
        try {
            Workbook workbook = new XSSFWorkbook(file.getInputStream());
            Sheet sheet = workbook.getSheetAt(0);

            for (Row row : sheet) {
                Cell iinCell = row.getCell(1);
                if (iinCell != null && iinCell.getCellType() == CellType.STRING) {
                    String plainIin = iinCell.getStringCellValue();
                    String encryptedIin = AesEncryptDecrypt.encrypt(plainIin);
                    if (encryptedIin != null) {
                        iinCell.setCellValue(encryptedIin);
                    }
                }
            }

            ByteArrayOutputStream out = new ByteArrayOutputStream();
            workbook.write(out);
            workbook.close();

            HttpHeaders headers = new HttpHeaders();
            headers.add(HttpHeaders.CONTENT_DISPOSITION, "attachment; filename=encrypted_iin.xlsx");

            return ResponseEntity.ok()
                    .headers(headers)
                    .contentType(MediaType.APPLICATION_OCTET_STREAM)
                    .body(out.toByteArray());
        } catch (Exception e) {
            e.printStackTrace();
            return ResponseEntity.badRequest().build(); // Return bad request in case of error
        }
    }

}