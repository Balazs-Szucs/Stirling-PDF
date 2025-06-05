package stirling.software.SPDF.controller.api.converters;

import java.io.IOException;
import java.nio.charset.StandardCharsets;

import org.jetbrains.annotations.NotNull;
import org.springframework.http.HttpHeaders;
import org.springframework.http.HttpStatus;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.ModelAttribute;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RestController;
import org.springframework.web.multipart.MultipartFile;

import io.swagger.v3.oas.annotations.Operation;
import io.swagger.v3.oas.annotations.tags.Tag;
import stirling.software.common.model.api.converters.EmlToPdfRequest;
import stirling.software.common.util.EmlToPdf;
import stirling.software.common.util.GeneralUtils;

@RestController
@RequestMapping("/api/v1/convert")
@Tag(name = "Convert", description = "Convert APIs")
public class ConvertEmlToPDF {

    @PostMapping(consumes = "multipart/form-data", value = "/eml/pdf")
    @Operation(
            summary = "Convert EML to PDF",
            description =
                    "This endpoint converts EML (email) files to PDF format with extensive"
                        + " customization options. Features include font settings, image constraints, display modes, attachment handling,"
                        + " and HTML debug output. Input: EML file, Output: PDF"
                        + " or HTML file. Type: SISO")
    public ResponseEntity<byte[]> convertEmlToPdf(@ModelAttribute EmlToPdfRequest request) {

        try {
            // Validate input
            MultipartFile inputFile = request.getFileInput();
            if (inputFile == null || inputFile.isEmpty()) {
                return ResponseEntity.badRequest()
                        .body("No file provided".getBytes(StandardCharsets.UTF_8));
            }

            // Validate file type - support both EML and MSG formats
            String originalFilename = inputFile.getOriginalFilename();
            if (originalFilename == null) {
                return ResponseEntity.badRequest()
                        .body("Please provide a valid filename".getBytes(StandardCharsets.UTF_8));
            }

            String lowerFilename = originalFilename.toLowerCase();
            boolean isEmlFile = lowerFilename.endsWith(".eml");

            if (!isEmlFile) {
                String supportedFormats = "EML files";
                return ResponseEntity.badRequest()
                        .body(("Please upload a valid " + supportedFormats)
                                .getBytes(StandardCharsets.UTF_8));
            }

            byte[] fileBytes = inputFile.getBytes();
            String baseFilename = GeneralUtils.convertToFileName(originalFilename);

            if (request.isDownloadHtml()) {
                try {
                    String htmlContent = EmlToPdf.convertEmlToHtml(fileBytes, originalFilename);

                    HttpHeaders headers = new HttpHeaders();
                    headers.setContentType(MediaType.TEXT_HTML);
                    headers.setContentDispositionFormData("attachment", baseFilename + ".html");
                    headers.add("X-Content-Type-Options", "nosniff");

                    return ResponseEntity.ok()
                            .headers(headers)
                            .body(htmlContent.getBytes(StandardCharsets.UTF_8));

                } catch (Exception e) {
                    return ResponseEntity.status(HttpStatus.INTERNAL_SERVER_ERROR)
                            .body(
                                    ("HTML conversion failed: " + e.getMessage())
                                            .getBytes(StandardCharsets.UTF_8));
                }
            }

            // Convert EML to PDF with enhanced options
            try {
                String weasyprintPath = "weasyprint";

                byte[] pdfBytes =
                        EmlToPdf.convertEmlToPdf(
                                weasyprintPath,
                                request,
                                fileBytes,
                                originalFilename,
                                false
                                );

                // Validate PDF output
                if (pdfBytes == null || pdfBytes.length == 0) {
                    return ResponseEntity.status(HttpStatus.INTERNAL_SERVER_ERROR)
                            .body(
                                    "PDF conversion failed - empty output"
                                            .getBytes(StandardCharsets.UTF_8));
                }

                // Create a response with appropriate headers
                HttpHeaders headers = new HttpHeaders();
                headers.setContentType(MediaType.APPLICATION_PDF);
                headers.setContentDispositionFormData("attachment", baseFilename + ".pdf");
                headers.add("X-Content-Type-Options", "nosniff");
                headers.setContentLength(pdfBytes.length);

                return ResponseEntity.ok().headers(headers).body(pdfBytes);

            } catch (InterruptedException e) {
                Thread.currentThread().interrupt();
                return ResponseEntity.status(HttpStatus.INTERNAL_SERVER_ERROR)
                        .body("Conversion was interrupted".getBytes(StandardCharsets.UTF_8));

            } catch (Exception e) {
                String errorMessage = getString(e);

                return ResponseEntity.status(HttpStatus.INTERNAL_SERVER_ERROR)
                        .body(errorMessage.getBytes(StandardCharsets.UTF_8));
            }

        } catch (IOException e) {
            return ResponseEntity.status(HttpStatus.INTERNAL_SERVER_ERROR)
                    .body("File processing error".getBytes(StandardCharsets.UTF_8));

        } catch (Exception e) {
            return ResponseEntity.status(HttpStatus.INTERNAL_SERVER_ERROR)
                    .body(
                            "An unexpected error occurred during conversion"
                                    .getBytes(StandardCharsets.UTF_8));
        }
    }

    private static @NotNull String getString(Exception e) {
        String errorMessage;
        if (e.getMessage() != null && e.getMessage().contains("Invalid EML")) {
            errorMessage =
                    "Invalid EML file format. Please ensure you've uploaded a valid email"
                            + " file.";
        } else if (e.getMessage() != null && e.getMessage().contains("WeasyPrint")) {
            errorMessage =
                    "PDF generation failed. This may be due to complex email formatting.";
        } else {
            errorMessage = "Conversion failed: " + e.getMessage();
        }
        return errorMessage;
    }
}
