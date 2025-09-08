package excelkarsilastir;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.util.*;

public class exceleslestir {
    public static void main(String[] args) {
        try {
            // Kodun çalıştığı dizini al
            File currentDir = new File(System.getProperty("user.dir"));

            // "ÇIKTI" klasörünü oluştur
            File outputDir = new File(currentDir, "ÇIKTI");
            if (!outputDir.exists()) {
                outputDir.mkdir();
            }

            // Dizindeki .csv ve .xlsx dosyalarını bul
            File[] files = currentDir.listFiles(new FilenameFilter() {
                public boolean accept(File dir, String name) {
                    return name.toLowerCase().endsWith(".csv") || name.toLowerCase().endsWith(".xlsx");
                }
            });

            if (files == null || files.length == 0) {
                System.err.println("Hata: Klasörde .csv veya .xlsx dosyası bulunamadı.");
                return;
            }

            File csvFile = null;
            File xlsxFile = null;

            for (File file : files) {
                if (file.getName().toLowerCase().endsWith(".csv")) {
                    if (csvFile != null) {
                        System.err.println("Hata: Klasörde birden fazla .csv dosyası var.");
                        return;
                    }
                    csvFile = file;
                } else if (file.getName().toLowerCase().endsWith(".xlsx")) {
                    if (xlsxFile != null) {
                        System.err.println("Hata: Klasörde birden fazla .xlsx dosyası var.");
                        return;
                    }
                    xlsxFile = file;
                }
            }

            if (csvFile == null || xlsxFile == null) {
                System.err.println("Hata: Hem bir .csv hem de bir .xlsx dosyası bulunamadı.");
                return;
            }

            // Sonuç dosyasının yolunu "ÇIKTI" klasörü olarak ayarla
            File sonucFile = new File(outputDir, "sonuc.xlsx");
            String sonucDosya = sonucFile.getAbsolutePath();

            Map<String, Integer> csvHeaderMap = new HashMap<>();
            List<String[]> csvRows = new ArrayList<>();

            try (BufferedReader br = new BufferedReader(new FileReader(csvFile))) {
                String line;
                boolean headerFound = false;
                while ((line = br.readLine()) != null) {
                    String[] values = line.split(";");
                    if (!headerFound) {
                        for (int i = 0; i < values.length; i++)
                            csvHeaderMap.put(normalize(values[i]), i);

                        if (!csvHeaderMap.containsKey("make") ||
                                !csvHeaderMap.containsKey("model") ||
                                !csvHeaderMap.containsKey("year")) {
                            System.err.println("CSV başlıkları bulunamadı! Mevcut başlıklar: " + Arrays.toString(values));
                            return;
                        }
                        headerFound = true;
                        continue;
                    }
                    csvRows.add(values);
                }
            }

            Workbook wb;
            try (FileInputStream xlsxStream = new FileInputStream(xlsxFile)) {
                wb = new XSSFWorkbook(xlsxStream);
            }
            Sheet sheet = wb.getSheetAt(0);

            Row xlsxHeaderRow = sheet.getRow(0);
            Map<String, Integer> xlsxHeaderMap = new HashMap<>();
            for (int i = 0; i < xlsxHeaderRow.getLastCellNum(); i++) {
                Cell c = xlsxHeaderRow.getCell(i);
                if (c != null) xlsxHeaderMap.put(normalize(c.toString()), i);
            }

            Integer markaIndex = xlsxHeaderMap.get("marka");
            Integer modelIndex = xlsxHeaderMap.get("model");
            Integer modelyiliIndex = xlsxHeaderMap.get("modelyili");
            Integer plakaIndex = xlsxHeaderMap.get("plaka");

            if (markaIndex == null || modelIndex == null || modelyiliIndex == null || plakaIndex == null) {
                System.err.println("XLSX başlıkları bulunamadı! Mevcut başlıklar: " + xlsxHeaderMap.keySet());
                wb.close();
                return;
            }

            Workbook wbSonuc = new XSSFWorkbook();
            Sheet sheetSonuc = wbSonuc.createSheet("Sonuc");
            int rowIndex = 0;

            Row headerRow = sheet.getRow(0);
            Row headerSonuc = sheetSonuc.createRow(rowIndex++);
            for (int i = 0; i < headerRow.getLastCellNum(); i++) {
                Cell oldCell = headerRow.getCell(i);
                Cell newCell = headerSonuc.createCell(i);
                if (oldCell != null) newCell.setCellValue(oldCell.toString());
            }
            headerSonuc.createCell(headerRow.getLastCellNum()).setCellValue("Kontrol_Sonucu");

            for (String[] csvRow : csvRows) {
                String csvMake = normalize(csvRow[csvHeaderMap.get("make")]);
                String csvModel = normalize(csvRow[csvHeaderMap.get("model")]);
                String csvYear = csvRow[csvHeaderMap.get("year")].trim();

                for (int r = 1; r <= sheet.getLastRowNum(); r++) {
                    Row xlsxRow = sheet.getRow(r);
                    if (xlsxRow == null) continue;

                    String xlsxMake = normalize(getCellValue(xlsxRow.getCell(markaIndex)));
                    String xlsxModel = normalize(getCellValue(xlsxRow.getCell(modelIndex)));
                    String xlsxYear = getCellValue(xlsxRow.getCell(modelyiliIndex)).trim();

                    if (!csvMake.equals(xlsxMake)) continue;
                    if (!hasPartialMatch(csvModel, xlsxModel)) continue;
                    if (!yearMatches(csvYear, xlsxYear)) continue;

                    Row newRow = sheetSonuc.createRow(rowIndex++);
                    for (int i = 0; i < xlsxRow.getLastCellNum(); i++) {
                        Cell oldCell = xlsxRow.getCell(i);
                        Cell newCell = newRow.createCell(i);
                        if (oldCell != null) newCell.setCellValue(oldCell.toString());
                    }

                    int ySayisi = 0;
                    List<String> ynColumns = Arrays.asList(
                            "Engine coolant temperature",
                            "Total distance",
                            "Fuel level",
                            "Engine speed",
                            "Speed",
                            "Total engine time"
                    );

                    List<String> kontrolList = new ArrayList<>();
                    for (String col : ynColumns) {
                        Integer idx = csvHeaderMap.get(normalize(col));
                        if (idx != null && idx < csvRow.length) kontrolList.add(csvRow[idx]);
                    }

                    for (String s : kontrolList) {
                        if (s != null && s.trim().equalsIgnoreCase("Y")) ySayisi++;
                    }
                    newRow.createCell(headerRow.getLastCellNum()).setCellValue(ySayisi > 3 ? "Y" : "N");
                }
            }

            Set<String> seenPlates = new HashSet<>();
            for (int r = 1; r <= sheetSonuc.getLastRowNum(); r++) {
                Row row = sheetSonuc.getRow(r);
                if (row == null) continue;

                Cell plakaCell = row.getCell(plakaIndex);
                if (plakaCell == null) continue;

                String plaka = plakaCell.toString().trim();
                if (seenPlates.contains(plaka)) {
                    sheetSonuc.removeRow(row);
                    if (r < sheetSonuc.getLastRowNum()) {
                        sheetSonuc.shiftRows(r + 1, sheetSonuc.getLastRowNum(), -1);
                    }
                    r--;
                } else {
                    seenPlates.add(plaka);
                }
            }

            try (FileOutputStream fos = new FileOutputStream(sonucDosya)) {
                wbSonuc.write(fos);
            }

            wb.close();
            wbSonuc.close();

            System.out.println("✅ İşlem tamamlandı. Sonuç dosyası: " + sonucDosya);

        } catch (Exception e) {
            e.printStackTrace();
        }
    }
    
    // Yardımcı metotlar aynı kalır
    private static String normalize(String s) {
        if (s == null) return "";
        return s.trim()
                .toLowerCase()
                .replaceAll("\\s+", "")
                .replaceAll("[\\[\\]\\(\\)\\.]", "")
                .replace("ç", "c")
                .replace("ş", "s")
                .replace("ı", "i")
                .replace("ü", "u")
                .replace("ö", "o")
                .replace("ğ", "g");
    }

    private static String getCellValue(Cell cell) {
        if (cell == null) return "";
        return cell.toString();
    }

    private static boolean yearMatches(String csvYear, String xlsxYear) {
        if (csvYear == null || csvYear.isEmpty() || xlsxYear == null || xlsxYear.isEmpty()) return false;
        int gYear;
        try {
            gYear = (int) Double.parseDouble(xlsxYear);
        } catch (NumberFormatException e) {
            return false;
        }

        csvYear = csvYear.trim();
        if (csvYear.matches("\\d{4}")) return gYear == Integer.parseInt(csvYear);
        else if (csvYear.matches("\\d{4}-\\d{4}")) {
            String[] parts = csvYear.split("-");
            int start = Integer.parseInt(parts[0]);
            int end = Integer.parseInt(parts[1]);
            return gYear >= start && gYear <= end;
        } else if (csvYear.matches("\\d{4}-")) {
            int start = Integer.parseInt(csvYear.substring(0, 4));
            return gYear >= start;
        } else return false;
    }

    private static boolean hasPartialMatch(String text1, String text2) {
        if (text1 == null || text2 == null) return false;
        String t1 = text1.trim().toLowerCase().replaceAll("[\\s\\-_/]", "");
        String t2 = text2.trim().toLowerCase().replaceAll("[\\s\\-_/]", "");
        return t1.contains(t2) || t2.contains(t1);
    }
}