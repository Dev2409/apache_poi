package com.dailycodework.excel2database.controller;

import com.dailycodework.excel2database.domain.Customer;
import com.dailycodework.excel2database.service.CustomerService;
import org.springframework.http.HttpStatus;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.*;
import org.springframework.web.multipart.MultipartFile;

import java.io.IOException;
import java.util.List;

@RestController
@RequestMapping("customers")
public class CustomerController {

    private final CustomerService customerService;

    public CustomerController(CustomerService customerService) {
        this.customerService = customerService;
    }

    @PostMapping("/upload-customers-data")
    public ResponseEntity<?> uploadCustomersData(@RequestParam("file") MultipartFile file) {
        try {
            // Parse the uploaded Excel file
            List<Customer> customers = customerService.parseExcelFile(file);

            // Save valid customers to the database
            customerService.saveCustomers(customers);

            return ResponseEntity.ok("Customers data uploaded and saved to database successfully");
        } catch (CustomerService.InvalidDataException e) {
            // Log the error
            System.err.println("Error parsing Excel file: " + e.getMessage());

            // Generate an Excel file with errors
            byte[] excelDataWithErrors;
            try {
                excelDataWithErrors = customerService.generateExcelWithErrors(file, e.getErrors());
            } catch (IOException ex) {
                return ResponseEntity.status(HttpStatus.INTERNAL_SERVER_ERROR)
                        .body("An error occurred while generating the error log.");
            }

            // Save the Excel file with errors to a file system or database
            // For simplicity, let's assume it's saved to the file system
            try {
                java.nio.file.Path path = java.nio.file.Paths.get("error_log_" + file.getOriginalFilename());
                java.nio.file.Files.write(path, excelDataWithErrors);
            } catch (IOException ex) {
                return ResponseEntity.status(HttpStatus.INTERNAL_SERVER_ERROR)
                        .body("An error occurred while saving the error log.");
            }

            return ResponseEntity.status(HttpStatus.BAD_REQUEST)
                    .body("Error parsing Excel file. Please check the file format.");
        } catch (Exception e) {
            // Handle other exceptions
            System.err.println("Error uploading customers data: " + e.getMessage());
            return ResponseEntity.status(HttpStatus.INTERNAL_SERVER_ERROR)
                    .body("An error occurred while processing the request.");
        }
    }
}
