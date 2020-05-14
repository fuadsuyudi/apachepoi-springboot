package net.suyudi.apachepoi.controller;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;

import javax.servlet.http.HttpServletRequest;

import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.core.io.UrlResource;
import org.springframework.http.HttpHeaders;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.bind.annotation.RestController;
import org.springframework.web.multipart.MultipartFile;

import net.suyudi.apachepoi.service.FileStorageService;
import net.suyudi.apachepoi.service.ReadingExcelService;
import net.suyudi.apachepoi.service.WritingExcelService;

import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.PathVariable;


/**
 * AuthorController
 */
@RestController
@RequestMapping("/authors")
public class AuthorController {

    @Autowired
    private FileStorageService fileStorageService;

    @Autowired
    private ReadingExcelService readingExcelService;

    @Autowired
    private WritingExcelService writingExcelService;

    @PostMapping("/import/stream")
    public String uploadFileStream(@RequestParam("file") MultipartFile file) {
        try {
            readingExcelService.GetData(file.getInputStream());
        } catch (Exception e) {
            return "Failed";
        }

        return "Success";
    }

    @PostMapping("/import")
    public String uploadFile(@RequestParam("file") MultipartFile file) {
        fileStorageService.storeFile(file);

        String filePath = fileStorageService.loadFileAsResource(file.getOriginalFilename()).toString();

        try {
            InputStream pathStream = new FileInputStream(new File(filePath));

            readingExcelService.GetData(pathStream);
        } catch (Exception e) {
            return "Failed: File path " + filePath;
        }

        return "Success: File path " + filePath;
    }
    
    @GetMapping("/export/{file}")
    public ResponseEntity<UrlResource> getMethodName(@PathVariable String file, HttpServletRequest request) {
        writingExcelService.buildFile(file);

        UrlResource resource = fileStorageService.getFileAsResource(file);

        String contentType = null;
        try {
            contentType = request.getServletContext().getMimeType(resource.getFile().getAbsolutePath());
        } catch (IOException ex) {
            System.out.println("Could not determine file type.");
        }

        if(contentType == null) {
            contentType = "application/octet-stream";
        }

        return ResponseEntity.ok()
                .contentType(MediaType.parseMediaType(contentType))
                .header(HttpHeaders.CONTENT_DISPOSITION, "attachment; filename=\"" + resource.getFilename() + "\"")
                .body(resource);
    }
    
}