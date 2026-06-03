package payu;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.text.SimpleDateFormat;
import java.util.*;

public class PayUReportGenerator {

    static class Stats {
        int count = 0;
        double amount = 0;

        void add(double val) {
            count++;
            amount += val;
        }

        String format() {
            return String.format("%d(%d)", count, (long) amount);
        }
    }

    private static final String[] PAYMENT_ORDER = {
            "UPI", "CC", "DC", "NB", "WALLET", "CARD"
    };

    public static void main(String[] args) {

        String filePath = "transaction_report_1295632776277478649.csv";

        Map<String, Map<String, Stats>> daily = new TreeMap<>();
        Map<String, Stats> grand = new HashMap<>();
        Map<String, Stats> dayTotal = new HashMap<>();
        Stats overall = new Stats();

        String minDate = null;
        String maxDate = null;

        try {

            if (filePath.endsWith(".xlsx")) {
                readExcel(filePath, daily, grand, dayTotal, overall);
            } else if (filePath.endsWith(".csv")) {
                readCSV(filePath, daily, grand, dayTotal, overall);
            } else {
                System.out.println("Unsupported file format!");
                return;
            }

            for (String d : daily.keySet()) {
                if (minDate == null || d.compareTo(minDate) < 0) minDate = d;
                if (maxDate == null || d.compareTo(maxDate) > 0) maxDate = d;
            }

        } catch (Exception e) {
            e.printStackTrace();
            return;
        }

        List<String[]> table = new ArrayList<>();

        String[] headerRow = new String[PAYMENT_ORDER.length + 2];
        headerRow[0] = "Date";
        System.arraycopy(PAYMENT_ORDER, 0, headerRow, 1, PAYMENT_ORDER.length);
        headerRow[headerRow.length - 1] = "Total";
        table.add(headerRow);

        try {
            SimpleDateFormat df = new SimpleDateFormat("yyyy-MM-dd");
            Date start = df.parse(minDate);
            Date end = df.parse(maxDate);

            Calendar cal = Calendar.getInstance();
            cal.setTime(start);

            while (!cal.getTime().after(end)) {

                String d = df.format(cal.getTime());

                Map<String, Stats> map = daily.getOrDefault(d, new HashMap<>());
                Stats t = dayTotal.getOrDefault(d, new Stats());

                String[] row = new String[PAYMENT_ORDER.length + 2];
                row[0] = d;

                for (int i = 0; i < PAYMENT_ORDER.length; i++) {
                    Stats s = map.get(PAYMENT_ORDER[i]);
                    row[i + 1] = (s != null) ? s.format() : "0(0)";
                }

                row[row.length - 1] = t.format();
                table.add(row);

                cal.add(Calendar.DATE, 1);
            }

        } catch (Exception e) {
            e.printStackTrace();
        }

        String[] g = new String[PAYMENT_ORDER.length + 2];
        g[0] = "GRAND TOTAL";

        for (int i = 0; i < PAYMENT_ORDER.length; i++) {
            g[i + 1] = grand.getOrDefault(PAYMENT_ORDER[i], new Stats()).format();
        }

        g[g.length - 1] = overall.format();
        table.add(g);

        String totalStr = overall.format();
        String[] totalRow = new String[PAYMENT_ORDER.length + 2];
        totalRow[0] = "TOTAL";
        Arrays.fill(totalRow, 1, totalRow.length, totalStr);
        table.add(totalRow);

        printTable(table);
    }

    // ================= CSV FIXED =================
    private static void readCSV(String filePath,
                                Map<String, Map<String, Stats>> daily,
                                Map<String, Stats> grand,
                                Map<String, Stats> dayTotal,
                                Stats overall) throws Exception {

        BufferedReader br = new BufferedReader(new FileReader(filePath));

        String headerLine = br.readLine();
        if (headerLine == null) return;

        List<String> headersList = parseCSVLine(headerLine);
        String[] headers = headersList.stream()
                .map(h -> h.toLowerCase().trim())
                .toArray(String[]::new);

        int dateCol = -1, modeCol = -1, amountCol = -1, statusCol = -1;

        for (int i = 0; i < headers.length; i++) {
            if (headers[i].equals("addedon")) dateCol = i;
            if (headers[i].equals("mode")) modeCol = i;
            if (headers[i].equals("amount")) amountCol = i;
            if (headers[i].equals("status")) statusCol = i;
        }

        String line;
        SimpleDateFormat inFmt = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss");
        SimpleDateFormat outFmt = new SimpleDateFormat("yyyy-MM-dd");

        while ((line = br.readLine()) != null) {

            List<String> values = parseCSVLine(line);
            if (values.size() <= Math.max(dateCol, Math.max(modeCol, amountCol))) continue;

            String status = values.get(statusCol).toUpperCase().trim();
            if (!status.contains("CAPTURED")) continue;

            String rawDate = values.get(dateCol).trim();
            String mode = values.get(modeCol).toUpperCase().trim();

            if (mode.contains("UPI")) mode = "UPI";
            else if (mode.contains("CREDIT")) mode = "CC";
            else if (mode.contains("DEBIT")) mode = "DC";
            else if (mode.contains("NET")) mode = "NB";
            else if (mode.contains("WALLET")) mode = "WALLET";

            double amount;
            try {
                amount = Double.parseDouble(values.get(amountCol).trim());
            } catch (Exception e) {
                continue;
            }

            String date;
            try {
                date = outFmt.format(inFmt.parse(rawDate));
            } catch (Exception e) {
                continue;
            }

            updateMaps(date, mode, amount, daily, grand, dayTotal, overall);
        }

        br.close();
    }

    // ================= SAFE CSV PARSER =================
    private static List<String> parseCSVLine(String line) {
        List<String> result = new ArrayList<>();
        StringBuilder current = new StringBuilder();
        boolean inQuotes = false;

        for (int i = 0; i < line.length(); i++) {
            char c = line.charAt(i);

            if (c == '"') {
                inQuotes = !inQuotes;
            } else if (c == ',' && !inQuotes) {
                result.add(current.toString());
                current.setLength(0);
            } else {
                current.append(c);
            }
        }

        result.add(current.toString());
        return result;
    }

    // ================= EXCEL FIXED =================
    private static void readExcel(String filePath,
                                  Map<String, Map<String, Stats>> daily,
                                  Map<String, Stats> grand,
                                  Map<String, Stats> dayTotal,
                                  Stats overall) throws Exception {

        FileInputStream fis = new FileInputStream(filePath);
        Workbook workbook = new XSSFWorkbook(fis);

        Sheet sheet = workbook.getSheetAt(0);
        Row header = sheet.getRow(0);

        int dateCol = -1, modeCol = -1, amountCol = -1, statusCol = -1;

        for (Cell cell : header) {
            String h = cell.getStringCellValue().toLowerCase().trim();

            if (h.equals("addedon")) dateCol = cell.getColumnIndex();
            if (h.equals("mode")) modeCol = cell.getColumnIndex();
            if (h.equals("amount")) amountCol = cell.getColumnIndex();
            if (h.equals("status")) statusCol = cell.getColumnIndex();
        }

        SimpleDateFormat inFmt = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss");
        SimpleDateFormat outFmt = new SimpleDateFormat("yyyy-MM-dd");

        for (int i = 1; i <= sheet.getLastRowNum(); i++) {

            Row row = sheet.getRow(i);
            if (row == null) continue;

            String status = getCellValue(row.getCell(statusCol)).toUpperCase().trim();
            if (!status.contains("CAPTURED")) continue;

            String rawDate = getCellValue(row.getCell(dateCol)).trim();
            String mode = getCellValue(row.getCell(modeCol)).toUpperCase().trim();

            if (mode.contains("UPI")) mode = "UPI";
            else if (mode.contains("CREDIT")) mode = "CC";
            else if (mode.contains("DEBIT")) mode = "DC";
            else if (mode.contains("NET")) mode = "NB";
            else if (mode.contains("WALLET")) mode = "WALLET";

            double amount;
            try {
                amount = Double.parseDouble(getCellValue(row.getCell(amountCol)).trim());
            } catch (Exception e) {
                continue;
            }

            String date;
            try {
                date = outFmt.format(inFmt.parse(rawDate));
            } catch (Exception e) {
                continue;
            }

            updateMaps(date, mode, amount, daily, grand, dayTotal, overall);
        }

        workbook.close();
    }

    private static void updateMaps(String date, String mode, double amount,
                                   Map<String, Map<String, Stats>> daily,
                                   Map<String, Stats> grand,
                                   Map<String, Stats> dayTotal,
                                   Stats overall) {

        daily.computeIfAbsent(date, k -> new HashMap<>())
                .computeIfAbsent(mode, k -> new Stats())
                .add(amount);

        dayTotal.computeIfAbsent(date, k -> new Stats()).add(amount);
        grand.computeIfAbsent(mode, k -> new Stats()).add(amount);
        overall.add(amount);
    }

    private static String getCellValue(Cell cell) {
        if (cell == null) return "";

        switch (cell.getCellType()) {
            case STRING:
                return cell.getStringCellValue();
            case NUMERIC:
                return String.valueOf(cell.getNumericCellValue());
            default:
                return "";
        }
    }

    private static void printTable(List<String[]> rows) {

        int cols = rows.get(0).length;

        StringBuilder line = new StringBuilder("+");
        for (int i = 0; i < cols; i++) {
            line.append("--------------+");
        }

        System.out.println(line);
        printRow(rows.get(0));
        System.out.println(line);

        for (int i = 1; i < rows.size(); i++) {
            printRow(rows.get(i));
            if (rows.get(i)[0].equals("GRAND TOTAL")) {
                System.out.println(line);
            }
        }

        System.out.println(line);
    }

    private static void printRow(String[] row) {
        System.out.print("|");

        for (int i = 0; i < row.length; i++) {
            if (i == 0) {
                System.out.printf(" %-12s |", row[i]);
            } else {
                System.out.printf(" %12s |", row[i]);
            }
        }

        System.out.println();
    }
}