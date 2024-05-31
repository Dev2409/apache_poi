package com.dailycodework.excel2database.service;

import com.dailycodework.excel2database.domain.Customer;
import com.dailycodework.excel2database.repository.CustomerRepository;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Service;
import org.springframework.web.multipart.MultipartFile;

import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.List;

@Service
public class CustomerService {

    private final CustomerRepository customerRepository;

    public CustomerService(CustomerRepository customerRepository) {
        this.customerRepository = customerRepository;
    }

    public void saveCustomersToDatabase(MultipartFile file) {
        try {
            List<Customer> validCustomers = parseExcelFile(file.getInputStream());
            saveValidCustomers(validCustomers);
        } catch (IOException e) {
            throw new RuntimeException("Failed to store excel data: " + e.getMessage());
        }
    }

    public List<Customer> getCustomers() {
        return customerRepository.findAll();
    }

    private void saveValidCustomers(List<Customer> validCustomers) {
        for (Customer customer : validCustomers) {
            try {
                customerRepository.save(customer);
            } catch (Exception e) {
                // Log error when saving customer
                System.err.println("Error saving customer: " + e.getMessage());
            }
        }
    }

    private List<Customer> parseExcelFile(InputStream is) throws IOException {
        List<Customer> validCustomers = new ArrayList<>();
        List<String> errors = new ArrayList<>();

        try (Workbook workbook = new XSSFWorkbook(is)) {
            Sheet sheet = workbook.getSheetAt(0);
            for (int i = 1; i <= sheet.getLastRowNum(); i++) {
                Row row = sheet.getRow(i);
                try {
                    Customer customer = validateAndParseRow(row, i);
                    validCustomers.add(customer);
                } catch (InvalidDataException e) {
                    errors.add("Row " + (i + 1) + ": " + e.getMessage());
                }
            }
        }

        // Log errors to a separate Excel sheet
        writeErrorLog(errors);

        return validCustomers;
    }

    private Customer validateAndParseRow(Row row, int rowNum) throws InvalidDataException {
        Customer customer = new Customer();
        try {
            customer.setCustomerId((int) row.getCell(0).getNumericCellValue());
            customer.setFirstName(row.getCell(1).getStringCellValue());
            customer.setLastName(row.getCell(2).getStringCellValue());
            customer.setCountry(row.getCell(3).getStringCellValue());
            customer.setTelephone((int) row.getCell(4).getNumericCellValue());
        } catch (Exception e) {
            throw new InvalidDataException("Invalid data type in row " + (rowNum));
        }
        return customer;
    }

    private void writeErrorLog(List<String> errors) {
        // Write error log to a separate Excel sheet
        // For simplicity, we print errors to console in this example
        for (String error : errors) {
            System.err.println("Error: " + error);
        }
    }

    static class InvalidDataException extends Exception {
        public InvalidDataException(String message) {
            super(message);
        }
    }
}
