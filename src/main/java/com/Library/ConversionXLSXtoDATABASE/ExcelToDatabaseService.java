package com.Library.ConversionXLSXtoDATABASE;

import jakarta.transaction.Transactional;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.jdbc.core.JdbcTemplate;
import org.springframework.stereotype.Service;

import java.io.File;
import java.io.FileInputStream;
import java.sql.SQLException;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;
@Service
public class ExcelToDatabaseService {

    private final JdbcTemplate jdbcTemplate;
    private static final String DATABASE_NAME = "conversiontest"; // Fixed main database

    public ExcelToDatabaseService(JdbcTemplate jdbcTemplate) {
        this.jdbcTemplate = jdbcTemplate;
    }

    @Transactional
    public void processExcelFile(String filePath) {
        try (FileInputStream fis = new FileInputStream(new File(filePath));
             Workbook workbook = new XSSFWorkbook(fis)) {

            jdbcTemplate.execute("USE " + DATABASE_NAME); // Use fixed database

            for (Sheet sheet : workbook) {
                String tableName = sheet.getSheetName().replaceAll("\\s+", "_").toLowerCase(); // Normalize table name
                System.out.println("Processing Table: " + tableName);

                List<String> columns = new ArrayList<>();
                List<List<String>> rows = new ArrayList<>();

                Iterator<Row> rowIterator = sheet.iterator();
                if (!rowIterator.hasNext()) continue; // Skip if no rows

                Row headerRow = rowIterator.next();
                for (Cell cell : headerRow) {
                    columns.add(cell.getStringCellValue().replaceAll("\\s+", "_").toLowerCase());
                }

                while (rowIterator.hasNext()) {
                    Row row = rowIterator.next();
                    List<String> rowData = new ArrayList<>();
                    for (int i = 0; i < columns.size(); i++) {
                        Cell cell = row.getCell(i, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
                        rowData.add(getCellValueAsString(cell));
                    }
                    rows.add(rowData);
                }

                createTable(tableName, columns);
                insertData(tableName, columns, rows);
            }
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    private void createTable(String tableName, List<String> columns) {
        // Remove empty column names to prevent SQL errors
        columns.removeIf(String::isEmpty);

        if (columns.isEmpty()) {
            System.out.println("Skipping table creation: No valid columns found in " + tableName);
            return;
        }

        StringBuilder createTableSQL = new StringBuilder("CREATE TABLE IF NOT EXISTS " + tableName + " (id INT AUTO_INCREMENT PRIMARY KEY, ");

        for (String column : columns) {
            createTableSQL.append(column).append(" VARCHAR(255), ");
        }

        // Remove the last comma
        createTableSQL.setLength(createTableSQL.length() - 2);
        createTableSQL.append(")");

        System.out.println("Executing Query: " + createTableSQL);
        jdbcTemplate.execute(createTableSQL.toString());
    }

    @Transactional
    private void insertData(String tableName, List<String> columns, List<List<String>> rows) {
        if (rows.isEmpty()) return;

        for (List<String> row : rows) {
            StringBuilder insertSQL = new StringBuilder("INSERT INTO " + tableName + " (");
            insertSQL.append(String.join(", ", columns)).append(") VALUES (");

            for (String value : row) {
                insertSQL.append("'").append(value.replace("'", "''")).append("', ");
            }
            insertSQL.setLength(insertSQL.length() - 2);
            insertSQL.append(")");

            System.out.println("Executing Insert Query: " + insertSQL);
            jdbcTemplate.execute(insertSQL.toString());
        }
    }

    private String getCellValueAsString(Cell cell) {
        switch (cell.getCellType()) {
            case STRING:
                return cell.getStringCellValue();
            case NUMERIC:
                return String.valueOf((long) cell.getNumericCellValue());
            case BOOLEAN:
                return String.valueOf(cell.getBooleanCellValue());
            default:
                return "";
        }
    }
}
