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
import java.util.*;

@Service
public class CustomerService {

    @Autowired
    private CustomerRepository customerRepository;

    @Autowired
    private ErrorLogRepository errorLogRepository;

    public List<Customer> parseExcelFile(MultipartFile file) throws InvalidDataException {
        List<Customer> customers = new ArrayList<>();
        List<ExcelError> errors = new ArrayList<>();
        Set<String> headerValues = new HashSet<>();

        try (InputStream inputStream = file.getInputStream()) {
            Workbook workbook = new XSSFWorkbook(inputStream);
            Sheet sheet = workbook.getSheetAt(0);

            // Get the header row
            Row headerRow = sheet.getRow(0);
            if (headerRow == null) {
                errors.add(new ExcelError(0, "Header row is missing"));
                throw new InvalidDataException(errors);
            }

            // Validate headers
            List<String> requiredHeaders = List.of("Customer ID", "First Name", "Last Name", "Country", "Telephone");
            Map<String, Integer> headerMap = new HashMap<>();
            for (int i = 0; i < headerRow.getLastCellNum(); i++) {
                Cell cell = headerRow.getCell(i);
                if (cell != null) {
                    String headerValue = cell.getStringCellValue().trim();
                    if (!headerValues.add(headerValue)) {
                        errors.add(new ExcelError(0, "Duplicate header value: " + headerValue));
                    }
                    headerMap.put(headerValue, i);
                }
            }

            for (String header : requiredHeaders) {
                if (!headerValues.contains(header)) {
                    errors.add(new ExcelError(0, "Missing required header: " + header));
                }
            }

            if (!errors.isEmpty()) {
                throw new InvalidDataException(errors);
            }

            // Process data rows
            for (int i = 1; i <= sheet.getLastRowNum(); i++) {
                Row row = sheet.getRow(i);
                List<ExcelError> rowErrors = new ArrayList<>();
                if (row != null) {
                    try {
                        Customer customer = new Customer();
                        for (String header : requiredHeaders) {
                            Cell cell = row.getCell(headerMap.get(header));
                            if (cell == null || cell.getCellType() == CellType.BLANK) {
                                rowErrors.add(new ExcelError(i, "Missing value for: " + header));
                                continue;
                            }

                            switch (header) {
                                case "CustomerID":
                                    if (cell.getCellType() != CellType.NUMERIC) {
                                        rowErrors.add(new ExcelError(i, "Invalid data type for customerId. Expected numeric."));
                                    } else {
                                        customer.setCustomerId((int) cell.getNumericCellValue());
                                    }
                                    break;
                                case "First Name":
                                    if (cell.getCellType() != CellType.STRING) {
                                        rowErrors.add(new ExcelError(i, "Invalid data type for firstName. Expected string."));
                                    } else {
                                        customer.setFirstName(cell.getStringCellValue());
                                    }
                                    break;
                                case "Last Name":
                                    if (cell.getCellType() != CellType.STRING) {
                                        rowErrors.add(new ExcelError(i, "Invalid data type for lastName. Expected string."));
                                    } else {
                                        customer.setLastName(cell.getStringCellValue());
                                    }
                                    break;
                                case "Country":
                                    if (cell.getCellType() != CellType.STRING) {
                                        rowErrors.add(new ExcelError(i, "Invalid data type for country. Expected string."));
                                    } else {
                                        customer.setCountry(cell.getStringCellValue());
                                    }
                                    break;
                                case "Telephone":
                                    if (cell.getCellType() != CellType.NUMERIC) {
                                        rowErrors.add(new ExcelError(i, "Invalid data type for telephone. Expected numeric."));
                                    } else {
                                        customer.setTelephone((int) cell.getNumericCellValue());
                                    }
                                    break;
                            }
                        }

                        if (rowErrors.isEmpty()) {
                            customers.add(customer);
                        } else {
                            errors.addAll(rowErrors);
                        }
                    } catch (Exception e) {
                        rowErrors.add(new ExcelError(i, e.getMessage()));
                    }
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
