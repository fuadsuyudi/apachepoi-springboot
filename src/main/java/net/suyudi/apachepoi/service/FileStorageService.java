package net.suyudi.apachepoi.service;

import org.springframework.stereotype.Service;
import org.springframework.web.multipart.MultipartFile;
import org.springframework.beans.factory.annotation.Value;
import org.springframework.core.io.UrlResource;

import net.suyudi.apachepoi.exception.FileStorageException;
import net.suyudi.apachepoi.exception.MyFileNotFoundException;

import java.io.IOException;
import java.net.MalformedURLException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.nio.file.StandardCopyOption;


@Service
public class FileStorageService {

    @Value("${local.path.file}")
    private String localPath;

    public String storeFile(MultipartFile file) {

        try {
            Path targetLocation = Paths.get(localPath).toAbsolutePath().normalize()
                    .resolve(file.getOriginalFilename());
            Files.copy(file.getInputStream(), targetLocation, StandardCopyOption.REPLACE_EXISTING);

            return file.getName();
        } catch (IOException ex) {
            throw new FileStorageException("Could not store file " + file.getName() + ". Please try again!", ex);
        }
    }

    public Path loadFileAsResource(String file) {
        try {
            Path filePath = Paths.get(localPath)
                    .toAbsolutePath()
                    .normalize()
                    .resolve(file)
                    .normalize();

            UrlResource resource = new UrlResource(filePath.toUri());

            if(resource.exists()) {
                return filePath;
            } else {
                throw new MyFileNotFoundException("File not found " + file);
            }
        } catch (MalformedURLException ex) {
            throw new MyFileNotFoundException("File not found " + file, ex);
        }
    }

    public UrlResource getFileAsResource(String file) {
        try {
            Path filePath = Paths.get(localPath)
                    .toAbsolutePath()
                    .normalize()
                    .resolve(file + ".xlsx")
                    .normalize();

            UrlResource resource = new UrlResource(filePath.toUri());

            if(resource.exists()) {
                return resource;
            } else {
                throw new MyFileNotFoundException("File not found " + file + ".xlsx");
            }
        } catch (MalformedURLException ex) {
            throw new MyFileNotFoundException("File not found " + file + ".xlsx", ex);
        }
    }
}
