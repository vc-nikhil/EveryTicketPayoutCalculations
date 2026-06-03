package com.orderreport;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.*;
import java.util.stream.Collectors;

public class OrderReportGenerator {

    static class Record {
        String date;
        String paymentMethod;
        double ticketPrice;

        Record(String date, String paymentMethod, double ticketPrice) {
            this.date = date;
            this.paymentMethod = paymentMethod;
            this.ticketPrice = ticketPrice;
        }
    }

    public static void main(String[] args) {
        String filePath = "C:\\Users\\Nikhil Sonawane\\eclipse-workspace\\OrderReport\\2026-05-01 to 2026-05-15_Order_History.xlsx";

        List<Record> records = readExcel(filePath);

        // Group by Date using TreeMap for sorted order
        Map<String, List<Record>> recordsByDate = new TreeMap<>(records.stream()
                .collect(Collectors.groupingBy(r -> r.date)));

        System.out.println("=== TICKET ANALYSIS REPORT ===");

        List<String> sortedDates = new ArrayList<>(recordsByDate.keySet());
        for (int i = 0; i < sortedDates.size(); i++) {
            if (i > 0) {
                System.out.println("========================================");
            }
            String date = sortedDates.get(i);
            System.out.println("Date: " + date);
            System.out.println("----------------------------------------");

            List<Record> recordsForDate = recordsByDate.get(date);

            // Group by Payment Method
            Map<String, List<Record>> recordsByPayment = recordsForDate.stream()
                    .collect(Collectors.groupingBy(r -> r.paymentMethod));

            // Sort payment methods alphabetically
            List<String> sortedPayments = recordsByPayment.keySet().stream().sorted().collect(Collectors.toList());

            // Print count and total ticket price
            for (String paymentMethod : sortedPayments) {
                List<Record> list = recordsByPayment.get(paymentMethod);
                double totalPrice = list.stream().mapToDouble(r -> r.ticketPrice).sum();

                System.out.println("Payment Method: " + paymentMethod);
                System.out.println("Total Records: " + list.size());
                System.out.println("Total Ticket Price: " + String.format("%.2f", totalPrice));
                System.out.println();
            }
        }

        // Summary Statistics
        System.out.println("=== SUMMARY STATISTICS ===");
        System.out.println("Total Records: " + records.size());
        double grandTotal = records.stream().mapToDouble(r -> r.ticketPrice).sum();
        System.out.println("Grand Total Amount: " + String.format("%.2f", grandTotal));

        // Payment Method Summary
        Map<String, List<Record>> byPayment = records.stream()
                .collect(Collectors.groupingBy(r -> r.paymentMethod));
        System.out.println("Payment Method Summary:");
        List<String> sortedAllPayments = byPayment.keySet().stream().sorted().collect(Collectors.toList());
        for (String pm : sortedAllPayments) {
            List<Record> list = byPayment.get(pm);
            int count = list.size();
            double total = list.stream().mapToDouble(r -> r.ticketPrice).sum();
            System.out.println("  " + pm + ": " + count + " records, Total: " + String.format("%.2f", total));
        }

        // Date Range
        if (!sortedDates.isEmpty()) {
            String minDate = sortedDates.get(0);
            String maxDate = sortedDates.get(sortedDates.size() - 1);
            System.out.println("Date Range: " + minDate + " to " + maxDate);
        }
    }

    private static List<Record> readExcel(String filePath) {
        List<Record> records = new ArrayList<>();

        try (FileInputStream fis = new FileInputStream(filePath);
             Workbook workbook = new XSSFWorkbook(fis)) {

            Sheet sheet = workbook.getSheetAt(0);
            Iterator<Row> iterator = sheet.iterator();
            iterator.next(); // Skip header row

            while (iterator.hasNext()) {
                Row row = iterator.next();

                // Skip empty or invalid rows
                if (row == null || row.getCell(8) == null || row.getCell(8).toString().trim().isEmpty()) {
                    continue;
                }

                // Read Date from "Purchased On" (column index 8)
                Cell dateCell = row.getCell(8);
                String dateStr = dateCell.getStringCellValue().trim();
                String date = dateStr.split("T")[0]; // Extract YYYY-MM-DD

                // Read Payment Method (column index 9)
                Cell paymentCell = row.getCell(9);
                String paymentMethod = (paymentCell != null) ? paymentCell.getStringCellValue().trim() : "";

                // Read Ticket Price (column index 5)
                double ticketPrice = 0.0;
                Cell priceCell = row.getCell(5);
                if (priceCell != null) {
                    if (priceCell.getCellType() == CellType.NUMERIC) {
                        ticketPrice = priceCell.getNumericCellValue();
                    } else if (priceCell.getCellType() == CellType.STRING) {
                        String priceText = priceCell.getStringCellValue().trim();
                        if (!priceText.isEmpty()) {
                            try {
                                ticketPrice = Double.parseDouble(priceText);
                            } catch (NumberFormatException e) {
                                System.err.println("Invalid ticket price at row " + row.getRowNum() + ": " + priceText);
                            }
                        }
                    }
                }

                records.add(new Record(date, paymentMethod, ticketPrice));
            }

        } catch (IOException e) {
            System.err.println("Error reading Excel file: " + e.getMessage());
        }

        return records;
    }
}
