package com.empujecomunitario.rest_server.controllers;

import com.empujecomunitario.rest_server.entity.Donation;
import com.empujecomunitario.rest_server.service.IDonationService;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.CrossOrigin;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RestController;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.util.List;

@RestController
@RequestMapping("/export")
@CrossOrigin(origins = "http://localhost:9091")
public class DonationExportController {

    private final IDonationService donationService;

    public DonationExportController(IDonationService donationService) {
        this.donationService = donationService;
    }

    @GetMapping("/donations")
    public ResponseEntity<byte[]> exportDonationsToExcel(@org.springframework.web.bind.annotation.RequestParam(value = "type", required = false) String type) throws IOException {
        try (Workbook workbook = new XSSFWorkbook()) {
            List<Donation> donations = donationService.findAll();

            // Si se solicita filtrar por tipo, ajustamos la lista:
            // type = "received" -> donaciones recibidas (madeByOurselves == false)
            // type = "made" -> donaciones realizadas por nosotros (madeByOurselves == true)
            if (type != null) {
                if (type.equalsIgnoreCase("received")) {
                    donations = donations.stream().filter(d -> d.getMadeByOurselves() != null && !d.getMadeByOurselves()).toList();
                } else if (type.equalsIgnoreCase("made")) {
                    donations = donations.stream().filter(d -> d.getMadeByOurselves() != null && d.getMadeByOurselves()).toList();
                }
            }

            for (Donation.Category category : Donation.Category.values()) {
                Sheet sheet = workbook.createSheet(category.name());

                Row header = sheet.createRow(0);
                header.createCell(0).setCellValue("Categoría");
                header.createCell(1).setCellValue("Descripción");
                header.createCell(2).setCellValue("Cantidad");
                header.createCell(3).setCellValue("Fecha última donación");
                header.createCell(4).setCellValue("Creado por nosotros");

                int rowIdx = 1;
                for (Donation donation : donations) {
                    if (donation.getIsDeleted()) continue; //ignoramos eliminadas
                    if (donation.getCategory() != category) continue;

                    Row row = sheet.createRow(rowIdx++);
                    row.createCell(0).setCellValue(donation.getCategory().name());
                    row.createCell(1).setCellValue(donation.getDescription() != null ? donation.getDescription() : "");
                    row.createCell(2).setCellValue(donation.getQuantity());
                    row.createCell(3).setCellValue(donation.getLastDonationDate().toString());
                    row.createCell(4).setCellValue(donation.getMadeByOurselves());
                }

                sheet.setColumnWidth(0, 4000);
                sheet.setColumnWidth(1, 10000);
                sheet.setColumnWidth(2, 4000);
                sheet.setColumnWidth(3, 6000);
                sheet.setColumnWidth(4, 5000);
            }

        ByteArrayOutputStream out = new ByteArrayOutputStream();
            workbook.write(out);
        String filename = "donaciones.xlsx";
        if (type != null) {
        if (type.equalsIgnoreCase("received")) filename = "donaciones_recibidas.xlsx";
        else if (type.equalsIgnoreCase("made")) filename = "donaciones_realizadas.xlsx";
        }

        return ResponseEntity.ok()
            .header("Content-Disposition", "attachment; filename=" + filename)
                    .contentType(MediaType.parseMediaType(
                            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"))
                    .body(out.toByteArray());
        }
    }
}