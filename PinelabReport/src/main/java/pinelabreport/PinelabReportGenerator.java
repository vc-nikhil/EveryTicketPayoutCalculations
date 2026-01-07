package pinelabreport;

import java.io.BufferedReader;
import java.io.FileReader;
import java.io.IOException;
import java.text.NumberFormat;
import java.util.*;
import java.util.stream.Collectors;

public class PinelabReportGenerator {

    static class Record {
        String date;          // column 23 (0-based 22)
        String paymentMethod; // column 1  (0-based 0)
        double ticketPrice;   // column 21 (0-based 20)

        Record(String date, String paymentMethod, double ticketPrice) {
            this.date = date;
            this.paymentMethod = paymentMethod;
            this.ticketPrice = ticketPrice;
        }
    }

    public static void main(String[] args) {
        // UPDATE THIS PATH TO YOUR FILE
        String filePath = "C:\\Users\\Nikhil Sonawane\\eclipse-workspace\\PinelabReport\\transactionReport (76).csv";

        try {
            List<Record> records = readCsv(filePath);

            if (records.isEmpty()) {
                System.out.println("No records found in the CSV file.");
                return;
            }

            // ---------- 1. Group by Date (preserve order) ----------
            Map<String, List<Record>> recordsByDate = records.stream()
                    .collect(Collectors.groupingBy(r -> r.date.split("T")[0], // keep only YYYY-MM-DD
                            LinkedHashMap::new, Collectors.toList()));

            long   grandCount = 0;
            double grandTotal = 0.0;
            NumberFormat nf = NumberFormat.getInstance();
            nf.setMinimumFractionDigits(2);
            nf.setMaximumFractionDigits(2);
            nf.setGroupingUsed(true);

            for (Map.Entry<String, List<Record>> dateEntry : recordsByDate.entrySet()) {
                String date = dateEntry.getKey();
                List<Record> dayRecords = dateEntry.getValue();

                System.out.println("\n===============================");
                System.out.println("Date: " + date);
                System.out.println("===============================");

                // ---------- 2. Group by Payment Method ----------
                Map<String, List<Record>> byMethod = dayRecords.stream()
                        .collect(Collectors.groupingBy(r -> r.paymentMethod,
                                LinkedHashMap::new, Collectors.toList()));

                long dayCount = 0;
                double dayTotal = 0.0;

                for (Map.Entry<String, List<Record>> methodEntry : byMethod.entrySet()) {
                    String method = methodEntry.getKey();
                    List<Record> list = methodEntry.getValue();

                    long   count = list.size();
                    double sum   = list.stream().mapToDouble(r -> r.ticketPrice).sum();

                    dayCount += count;
                    dayTotal += sum;

                    System.out.printf("Payment Method: %-12s | Count: %4d | Amount: %12s%n",
                            method, count, nf.format(sum));
                }

                // ---------- 3. Day total ----------
                System.out.println("-------------------------------------------------");
                System.out.printf("DAY TOTAL                | Count: %4d | Amount: %12s%n",
                        dayCount, nf.format(dayTotal));
                System.out.println("-------------------------------------------------");

                grandCount += dayCount;
                grandTotal += dayTotal;
            }

            // ---------- 4. Grand total ----------
            System.out.println("\n===============================");
            System.out.printf("GRAND TOTAL (all days)   | Count: %4d | Amount: %12s%n",
                    grandCount, nf.format(grandTotal));
            System.out.println("===============================");

        } catch (Exception e) {
            System.err.println("Error: " + e.getMessage());
            e.printStackTrace();
        }
    }

    /* -------------------------------------------------------------
       CSV READER – unchanged (handles quoted fields, commas inside quotes)
       ------------------------------------------------------------- */
    private static List<Record> readCsv(String filePath) throws IOException {
        List<Record> records = new ArrayList<>();
        boolean isFirstLine = true;

        try (BufferedReader br = new BufferedReader(new FileReader(filePath))) {
            String line;
            while ((line = br.readLine()) != null) {
                if (isFirstLine) {
                    isFirstLine = false;
                    continue;               // skip header row
                }

                String[] values = parseCsvLine(line);
                if (values.length < 23) {   // need at least column 23 (date)
                    System.err.println("Skipping short row (less than 23 cols): " + line);
                    continue;
                }

                String method   = values[0].trim();                     // "UPI" or "Card"
                String priceStr = values[20].trim().replace(",", ""); // price column
                String dateStr  = values[22].trim();                  // transaction date

                double price;
                try {
                    price = Double.parseDouble(priceStr);
                } catch (NumberFormatException e) {
                    System.err.println("Invalid price in row (price='" + values[20] + "'): " + line);
                    continue;
                }

                records.add(new Record(dateStr, method, price));
            }
        }
        return records;
    }

    /** Simple CSV splitter that respects double-quotes */
    private static String[] parseCsvLine(String line) {
        List<String> result = new ArrayList<>();
        StringBuilder field = new StringBuilder();
        boolean inQuotes = false;

        for (int i = 0; i < line.length(); i++) {
            char c = line.charAt(i);

            if (c == '"') {
                inQuotes = !inQuotes;
            } else if (c == ',' && !inQuotes) {
                result.add(field.toString());
                field = new StringBuilder();
            } else {
                field.append(c);
            }
        }
        result.add(field.toString());   // last field
        return result.toArray(new String[0]);
    }
}