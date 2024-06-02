package com.dailycodework.excel2database.service;

import com.dailycodework.excel2database.domain.Customer;
import com.dailycodework.excel2database.domain.ErrorLog;
import com.dailycodework.excel2database.repository.CustomerRepository;
import com.dailycodework.excel2database.repository.ErrorLogRepository;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;
import org.springframework.web.multipart.MultipartFile;

import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.List;

@Service
public class CustomerService {

    @Autowired
    private CustomerRepository customerRepository;

    @Autowired
    private ErrorLogRepository errorLogRepository;

    public List<Customer> parseExcelFile(MultipartFile file) throws InvalidDataException {
        List<Customer> customers = new ArrayList<>();
        List<ExcelError> errors = new ArrayList<>();

        try (InputStream inputStream = file.getInputStream()) {
            Workbook workbook = new XSSFWorkbook(inputStream);
            Sheet sheet = workbook.getSheetAt(0);
            for (Row row : sheet) {
                if (row.getRowNum() == 0) {
                    continue; // Skip header row
                }
                try {
                    Cell customerIdCell = row.getCell(0);
                    Cell firstNameCell = row.getCell(1);
                    Cell lastNameCell = row.getCell(2);
                    Cell countryCell = row.getCell(3);
                    Cell telephoneCell = row.getCell(4);

                    Integer customerId = (int) customerIdCell.getNumericCellValue();
                    String firstName = firstNameCell != null ? firstNameCell.getStringCellValue() : null;
                    String lastName = lastNameCell != null ? lastNameCell.getStringCellValue() : null;
                    String country = countryCell != null ? countryCell.getStringCellValue() : null;
                    Integer telephone = (int) telephoneCell.getNumericCellValue();

                    if (firstName == null || firstName.trim().isEmpty()) {
                        throw new IllegalArgumentException("First name is missing");
                    }
                    if (lastName == null || lastName.trim().isEmpty()) {
                        throw new IllegalArgumentException("Last name is missing");
                    }
                    if (country == null || country.trim().isEmpty()) {
                        throw new IllegalArgumentException("Country is missing");
                    }
                    if (telephone == null) {
                        throw new IllegalArgumentException("Telephone is missing");
                    }

                    customers.add(new Customer(customerId, firstName, lastName, country, telephone));
                } catch (Exception e) {
                    errors.add(new ExcelError(row.getRowNum(), e.getMessage()));
                }
            }

            if (!errors.isEmpty()) {
                throw new InvalidDataException(errors);
            }
        } catch (IOException e) {
            throw new InvalidDataException("Error reading Excel file: " + e.getMessage());
        }

        return customers;
    }

    public void saveCustomers(List<Customer> customers) {
        customerRepository.saveAll(customers);
    }

    public byte[] generateExcelWithErrors(MultipartFile file, List<ExcelError> errors) throws IOException {
        Workbook workbook = new XSSFWorkbook(file.getInputStream());
        Sheet sheet = workbook.getSheetAt(0);

        // Add an error column header
        Row headerRow = sheet.getRow(0);
        Cell headerCell = headerRow.createCell(headerRow.getLastCellNum());
        headerCell.setCellValue("Error");

        // Add error messages to the corresponding rows
        for (ExcelError error : errors) {
            Row row = sheet.getRow(error.getRow());
            if (row != null) {
                Cell errorCell = row.createCell(row.getLastCellNum());
                errorCell.setCellValue(error.getMessage());
            }
        }

        // Write the workbook to a byte array
        ByteArrayOutputStream outputStream = new ByteArrayOutputStream();
        workbook.write(outputStream);
        workbook.close();

        return outputStream.toByteArray();
    }

    public void saveErrorLog(String filename, String errorMessage) {
        ErrorLog errorLog = new ErrorLog(errorMessage, filename);
        errorLogRepository.save(errorLog);
    }

    public static class InvalidDataException extends Exception {
        private final List<ExcelError> errors;

        public InvalidDataException(List<ExcelError> errors) {
            super("Invalid data in Excel file");
            this.errors = errors;
        }

        public InvalidDataException(String message) {
            super(message);
            this.errors = new ArrayList<>();
            this.errors.add(new ExcelError(-1, message));
        }

        public List<ExcelError> getErrors() {
            return errors;
        }
    }

    public static class ExcelError {
        private final int row;
        private final String message;

        public ExcelError(int row, String message) {
            this.row = row;
            this.message = message;
        }

        public int getRow() {
            return row;
        }

        public String getMessage() {
            return message;
        }
    }
}
