package stirling.software.common.util;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.nio.charset.StandardCharsets;
import java.util.ArrayList;
import java.util.Date;
import java.util.GregorianCalendar;
import java.util.HashMap;
import java.util.List;
import java.util.Locale;
import java.util.Map;
import java.util.Properties;

import org.apache.pdfbox.Loader;
import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.pdmodel.PDDocumentNameDictionary;
import org.apache.pdfbox.pdmodel.PDEmbeddedFilesNameTreeNode;
import org.apache.pdfbox.pdmodel.PDPage;
import org.apache.pdfbox.pdmodel.common.PDRectangle;
import org.apache.pdfbox.pdmodel.common.filespecification.PDComplexFileSpecification;
import org.apache.pdfbox.pdmodel.common.filespecification.PDEmbeddedFile;
import org.apache.pdfbox.pdmodel.interactive.annotation.PDAnnotationFileAttachment;
import org.apache.pdfbox.text.PDFTextStripper;
import org.apache.pdfbox.text.TextPosition;

import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.NoArgsConstructor;
import lombok.experimental.UtilityClass;
import lombok.extern.slf4j.Slf4j;

import stirling.software.common.model.api.converters.EmlToPdfRequest;

/**
 * Enhanced import stirling.software.common.model.api.converters.EmlToPdfRequest; y handles missing
 * Jakarta Mail dependencies.
 */
@Slf4j
@UtilityClass
public class EmlToPdf {

    // Jakarta Mail availability check
    private static Boolean jakartaMailAvailable = null;

    /** Check if Jakarta Mail is available in the classpath */
    private static boolean isJakartaMailAvailable() {
        if (jakartaMailAvailable == null) {
            try {
                Class.forName("jakarta.mail.internet.MimeMessage");
                Class.forName("jakarta.mail.Session");
                jakartaMailAvailable = true;
            } catch (ClassNotFoundException e) {
                jakartaMailAvailable = false;
            }
        }
        return jakartaMailAvailable;
    }

    /** Converts an EML file to HTML format with enhanced error handling. */
    public static String convertEmlToHtml(byte[] emlBytes, String fileName) throws IOException {

        if (emlBytes == null || emlBytes.length == 0) {
            throw new IllegalArgumentException("EML file is empty or null");
        }

        if (!isValidEmlFormat(emlBytes)) {
            throw new IllegalArgumentException("Invalid EML file format");
        }

        if (isJakartaMailAvailable()) {
            return convertEmlToHtmlAdvanced(emlBytes, fileName, null);
        } else {
            return convertEmlToHtmlBasic(emlBytes, fileName);
        }
    }

    /** Enhanced EML to PDF conversion with comprehensive configuration options. */
    public static byte[] convertEmlToPdf(
            String weasyprintPath,
            EmlToPdfRequest request,
            byte[] emlBytes,
            String fileName,
            boolean disableSanitize)
            throws IOException, InterruptedException {

        if (emlBytes == null || emlBytes.length == 0) {
            throw new IllegalArgumentException("EML file is empty or null");
        }

        if (!isValidEmlFormat(emlBytes)) {
            throw new IllegalArgumentException("Invalid EML file format");
        }

        // Generate HTML representation
        String htmlContent;
        EmailContent emailContent = null;
        if (isJakartaMailAvailable()) {
            emailContent = extractEmailContentAdvanced(emlBytes, fileName, request);
            htmlContent = generateEnhancedEmailHtml(emailContent, request);
        } else {
            htmlContent = convertEmlToHtmlBasic(emlBytes, fileName, request);
        }

        // Create enhanced HTML to PDF request
        stirling.software.common.model.api.converters.HTMLToPdfRequest htmlRequest =
                createHtmlRequest(request);

        // Convert HTML to PDF first (without attachments in the HTML)
        byte[] pdfBytes;
        try {
            pdfBytes =
                    FileToPdf.convertHtmlToPdf(
                            weasyprintPath,
                            htmlRequest,
                            htmlContent.getBytes(StandardCharsets.UTF_8),
                            "email.html",
                            disableSanitize,
                            request);
        } catch (Exception e) {
            // Try with simplified HTML
            String simplifiedHtml = htmlContent.replaceAll("(?i)<script[^>]*>.*?</script>", "");
            simplifiedHtml = simplifiedHtml.replaceAll("(?i)<style[^>]*>.*?</style>", "");

            pdfBytes =
                    FileToPdf.convertHtmlToPdf(
                            weasyprintPath,
                            htmlRequest,
                            simplifiedHtml.getBytes(StandardCharsets.UTF_8),
                            "email.html",
                            disableSanitize,
                            request);
        }

        // Apply attachments to the PDF if available and requested
        if (emailContent != null
                && request != null
                && request.isIncludeAttachments()
                && !emailContent.attachments.isEmpty()) {

            try {
                pdfBytes = attachFilesToPdf(pdfBytes, emailContent.attachments);
            } catch (Exception e) {
                // Continue with PDF without attachments rather than failing completely
            }
        }

        return pdfBytes;
    }

    /** Basic EML to HTML conversion without Jakarta Mail dependencies */
    private static String convertEmlToHtmlBasic(byte[] emlBytes, String fileName) {
        return convertEmlToHtmlBasic(emlBytes, fileName, null);
    }

    private static String convertEmlToHtmlBasic(
            byte[] emlBytes, String fileName, EmlToPdfRequest request) {
        if (emlBytes == null || emlBytes.length == 0) {
            throw new IllegalArgumentException("EML file is empty or null");
        }

        String emlContent = new String(emlBytes, StandardCharsets.UTF_8);

        // Basic email parsing
        String subject = extractBasicHeader(emlContent, "Subject:");
        String from = extractBasicHeader(emlContent, "From:");
        String to = extractBasicHeader(emlContent, "To:");
        String cc = extractBasicHeader(emlContent, "Cc:");
        String bcc = extractBasicHeader(emlContent, "Bcc:");
        String date = extractBasicHeader(emlContent, "Date:");

        // Try to extract HTML content
        String htmlBody = extractHtmlBody(emlContent);
        if (htmlBody == null) {
            String textBody = extractTextBody(emlContent);
            htmlBody =
                    convertTextToHtml(
                            textBody != null ? textBody : "Email content could not be parsed");
        }

        // Generate HTML with custom styling based on request
        StringBuilder html = new StringBuilder();
        html.append("<!DOCTYPE html>\n");
        html.append("<html><head><meta charset=\"UTF-8\">\n");
        html.append("<title>").append(escapeHtml(subject)).append("</title>\n");
        html.append("<style>\n");
        appendEnhancedStyles(html, request);
        html.append("</style>\n");
        html.append("</head><body>\n");

        html.append("<div class=\"email-container\">\n");
        html.append("<div class=\"email-header\">\n");
        html.append("<h1>").append(escapeHtml(subject)).append("</h1>\n");
        html.append("<div class=\"email-meta\">\n");
        html.append("<div><strong>From:</strong> ").append(escapeHtml(from)).append("</div>\n");
        html.append("<div><strong>To:</strong> ").append(escapeHtml(to)).append("</div>\n");

        // Include CC and BCC if present and requested
        if (request != null && request.isIncludeAllRecipients()) {
            if (cc != null && !cc.trim().isEmpty()) {
                html.append("<div><strong>CC:</strong> ").append(escapeHtml(cc)).append("</div>\n");
            }
            if (bcc != null && !bcc.trim().isEmpty()) {
                html.append("<div><strong>BCC:</strong> ")
                        .append(escapeHtml(bcc))
                        .append("</div>\n");
            }
        }

        if (date != null && !date.trim().isEmpty()) {
            html.append("<div><strong>Date:</strong> ").append(escapeHtml(date)).append("</div>\n");
        }
        html.append("</div></div>\n");

        html.append("<div class=\"email-body\">\n");
        html.append(processEmailHtmlBody(htmlBody, request));
        html.append("</div>\n");

        // Add attachment information - always check for and display attachments
        String attachmentInfo = extractAttachmentInfo(emlContent);
        if (!attachmentInfo.isEmpty()) {
            html.append("<div class=\"attachment-section\">\n");
            html.append("<h3>Attachments</h3>\n");
            html.append(attachmentInfo);

            // Add status message about attachment inclusion
            if (request != null && request.isIncludeAttachments()) {
                html.append("<div class=\"attachment-inclusion-note\">\n");
                html.append(
                        "<p><strong>Note:</strong> Attachments are saved as external files and linked in this PDF. Click the links to open files externally.</p>\n");
                html.append("</div>\n");
            } else {
                html.append("<div class=\"attachment-info-note\">\n");
                html.append(
                        "<p><em>Attachment information displayed - files not included in PDF. Enable 'Include attachments' to embed files.</em></p>\n");
                html.append("</div>\n");
            }

            html.append("</div>\n");
        }

        // Show advanced features status if requested
        if (request != null) {
            html.append("<div class=\"advanced-features-notice\">\n");
            html.append(
                    "<p><em>Note: Some advanced features require Jakarta Mail dependencies.</em></p>\n");
            html.append("</div>\n");
        }

        html.append("</div>\n");
        html.append("</body></html>");

        return html.toString();
    }

    /** Advanced EML to HTML conversion using Jakarta Mail */
    private static EmailContent extractEmailContentAdvanced(
            byte[] emlBytes, String fileName, EmlToPdfRequest request) {
        try {
            // Use Jakarta Mail for advanced processing
            Class<?> sessionClass = Class.forName("jakarta.mail.Session");
            Class<?> mimeMessageClass = Class.forName("jakarta.mail.internet.MimeMessage");

            java.lang.reflect.Method getDefaultInstance =
                    sessionClass.getMethod("getDefaultInstance", Properties.class);
            Object session = getDefaultInstance.invoke(null, new Properties());

            java.lang.reflect.Constructor<?> mimeMessageConstructor =
                    mimeMessageClass.getConstructor(sessionClass, java.io.InputStream.class);
            Object message =
                    mimeMessageConstructor.newInstance(session, new ByteArrayInputStream(emlBytes));

            // Extract content using reflection
            EmailContent content = extractEmailContentAdvanced(message, request);

            return content;

        } catch (Exception e) {
            // Create basic EmailContent from basic processing
            EmailContent content = new EmailContent();
            String htmlContent = convertEmlToHtmlBasic(emlBytes, fileName, request);
            content.htmlBody = htmlContent;
            return content;
        }
    }

    private static String convertEmlToHtmlAdvanced(
            byte[] emlBytes, String fileName, EmlToPdfRequest request) {
        EmailContent content = extractEmailContentAdvanced(emlBytes, fileName, request);
        return generateEnhancedEmailHtml(content, request);
    }

    // Utility methods

    private static String extractAttachmentInfo(String emlContent) {
        StringBuilder attachmentInfo = new StringBuilder();
        try {
            String[] lines = emlContent.split("\r?\n");
            boolean inHeaders = true;
            String currentContentType = "";
            String currentDisposition = "";
            String currentFilename = "";
            String currentEncoding = "";
            boolean inMultipart = false;
            String boundary = "";

            // First pass: find boundary for multipart messages
            for (String line : lines) {
                String lowerLine = line.toLowerCase().trim();
                if (lowerLine.startsWith("content-type:") && lowerLine.contains("multipart")) {
                    if (lowerLine.contains("boundary=")) {
                        int boundaryStart = lowerLine.indexOf("boundary=") + 9;
                        String boundaryPart = line.substring(boundaryStart).trim();
                        if (boundaryPart.startsWith("\"")) {
                            boundary = boundaryPart.substring(1, boundaryPart.indexOf("\"", 1));
                        } else {
                            int spaceIndex = boundaryPart.indexOf(" ");
                            boundary =
                                    spaceIndex > 0
                                            ? boundaryPart.substring(0, spaceIndex)
                                            : boundaryPart;
                        }
                        inMultipart = true;
                        break;
                    }
                }
                if (line.trim().isEmpty()) break;
            }

            // Second pass: extract attachment information
            inHeaders = true;
            for (String line : lines) {
                String lowerLine = line.toLowerCase().trim();

                // Check for boundary markers in multipart messages
                if (inMultipart && line.trim().startsWith("--" + boundary)) {
                    // Reset for new part
                    currentContentType = "";
                    currentDisposition = "";
                    currentFilename = "";
                    currentEncoding = "";
                    inHeaders = true;
                    continue;
                }

                if (inHeaders && line.trim().isEmpty()) {
                    inHeaders = false;

                    // Process accumulated attachment info
                    if (isAttachment(currentDisposition, currentFilename, currentContentType)) {
                        addAttachmentToInfo(
                                attachmentInfo,
                                currentFilename,
                                currentContentType,
                                currentEncoding);

                        // Reset for next attachment
                        currentContentType = "";
                        currentDisposition = "";
                        currentFilename = "";
                        currentEncoding = "";
                    }
                    continue;
                }

                if (!inHeaders) continue; // Skip body content

                // Parse headers
                if (lowerLine.startsWith("content-type:")) {
                    currentContentType = line.substring(13).trim();
                } else if (lowerLine.startsWith("content-disposition:")) {
                    currentDisposition = line.substring(20).trim();
                    // Extract filename if present
                    currentFilename = extractFilenameFromDisposition(currentDisposition);
                } else if (lowerLine.startsWith("content-transfer-encoding:")) {
                    currentEncoding = line.substring(26).trim();
                } else if (line.startsWith(" ") || line.startsWith("\t")) {
                    // Continuation of previous header
                    if (!currentDisposition.isEmpty() && currentDisposition.contains("filename=")) {
                        currentDisposition += " " + line.trim();
                        currentFilename = extractFilenameFromDisposition(currentDisposition);
                    } else if (!currentContentType.isEmpty()) {
                        currentContentType += " " + line.trim();
                    }
                }
            }

            // Don't forget the last attachment if we ended while still in headers
            if (isAttachment(currentDisposition, currentFilename, currentContentType)) {
                addAttachmentToInfo(
                        attachmentInfo, currentFilename, currentContentType, currentEncoding);
            }

        } catch (Exception e) {
        }
        return attachmentInfo.toString();
    }

    private static boolean isAttachment(String disposition, String filename, String contentType) {
        return (disposition.toLowerCase().contains("attachment") && !filename.isEmpty())
                || (!filename.isEmpty() && !contentType.toLowerCase().startsWith("text/"))
                || (contentType.toLowerCase().contains("application/") && !filename.isEmpty());
    }

    private static String extractFilenameFromDisposition(String disposition) {
        if (disposition.contains("filename=")) {
            int filenameStart = disposition.toLowerCase().indexOf("filename=") + 9;
            int filenameEnd = disposition.indexOf(";", filenameStart);
            if (filenameEnd == -1) filenameEnd = disposition.length();
            String filename = disposition.substring(filenameStart, filenameEnd).trim();
            return filename.replaceAll("^\"|\"$", "");
        }
        return "";
    }

    private static void addAttachmentToInfo(
            StringBuilder attachmentInfo, String filename, String contentType, String encoding) {
        attachmentInfo.append("<div class=\"attachment-item\">\n");

        // Create attachment info without paperclip emoji
        String uniqueId = "attachment_" + filename.hashCode() + "_" + System.nanoTime();
        attachmentInfo
                .append("<span class=\"attachment-link\" id=\"")
                .append(uniqueId)
                .append("\" data-filename=\"")
                .append(escapeHtml(filename))
                .append("\">\n")
                .append("<span class=\"attachment-name\">")
                .append(escapeHtml(filename))
                .append("</span>\n");

        // Add content type and encoding info
        if (!contentType.isEmpty() || !encoding.isEmpty()) {
            attachmentInfo.append("<span class=\"attachment-type\"> (");
            if (!contentType.isEmpty()) {
                attachmentInfo.append(escapeHtml(contentType));
            }
            if (!encoding.isEmpty()) {
                if (!contentType.isEmpty()) attachmentInfo.append(", ");
                attachmentInfo.append("encoding: ").append(escapeHtml(encoding));
            }
            attachmentInfo.append(")</span>\n");
        }

        attachmentInfo.append("</span>\n");
        attachmentInfo.append("</div>\n");
    }

    private static boolean isValidEmlFormat(byte[] emlBytes) {
        try {
            int checkLength = Math.min(emlBytes.length, 8192);
            String content = new String(emlBytes, 0, checkLength, StandardCharsets.UTF_8);
            String lowerContent = content.toLowerCase();

            boolean hasFrom =
                    lowerContent.contains("from:") || lowerContent.contains("return-path:");
            boolean hasSubject = lowerContent.contains("subject:");
            boolean hasMessageId = lowerContent.contains("message-id:");
            boolean hasDate = lowerContent.contains("date:");
            boolean hasTo =
                    lowerContent.contains("to:")
                            || lowerContent.contains("cc:")
                            || lowerContent.contains("bcc:");
            boolean hasMimeStructure =
                    lowerContent.contains("multipart/")
                            || lowerContent.contains("text/plain")
                            || lowerContent.contains("text/html")
                            || lowerContent.contains("boundary=");

            int headerCount = 0;
            if (hasFrom) headerCount++;
            if (hasSubject) headerCount++;
            if (hasMessageId) headerCount++;
            if (hasDate) headerCount++;
            if (hasTo) headerCount++;

            return headerCount >= 2 || hasMimeStructure;

        } catch (Exception e) {
            return false;
        }
    }

    private static String extractBasicHeader(String emlContent, String headerName) {
        try {
            String[] lines = emlContent.split("\r?\n");
            for (int i = 0; i < lines.length; i++) {
                String line = lines[i];
                if (line.toLowerCase().startsWith(headerName.toLowerCase())) {
                    StringBuilder value =
                            new StringBuilder(line.substring(headerName.length()).trim());
                    // Handle multi-line headers
                    for (int j = i + 1; j < lines.length; j++) {
                        if (lines[j].startsWith(" ") || lines[j].startsWith("\t")) {
                            value.append(" ").append(lines[j].trim());
                        } else {
                            break;
                        }
                    }
                    return value.toString();
                }
                if (line.trim().isEmpty()) break;
            }
        } catch (Exception e) {
        }
        return "";
    }

    private static String extractHtmlBody(String emlContent) {
        try {
            String lowerContent = emlContent.toLowerCase();
            int htmlStart = lowerContent.indexOf("content-type: text/html");
            if (htmlStart == -1) return null;

            int bodyStart = emlContent.indexOf("\r\n\r\n", htmlStart);
            if (bodyStart == -1) bodyStart = emlContent.indexOf("\n\n", htmlStart);
            if (bodyStart == -1) return null;

            bodyStart += (emlContent.charAt(bodyStart + 1) == '\r') ? 4 : 2;
            int bodyEnd = findPartEnd(emlContent, bodyStart);

            return emlContent.substring(bodyStart, bodyEnd).trim();

        } catch (Exception e) {
            return null;
        }
    }

    private static String extractTextBody(String emlContent) {
        try {
            String lowerContent = emlContent.toLowerCase();
            int textStart = lowerContent.indexOf("content-type: text/plain");
            if (textStart == -1) {
                int bodyStart = emlContent.indexOf("\r\n\r\n");
                if (bodyStart == -1) bodyStart = emlContent.indexOf("\n\n");
                if (bodyStart != -1) {
                    bodyStart += (emlContent.charAt(bodyStart + 1) == '\r') ? 4 : 2;
                    int bodyEnd = findPartEnd(emlContent, bodyStart);
                    return emlContent.substring(bodyStart, bodyEnd).trim();
                }
                return null;
            }

            int bodyStart = emlContent.indexOf("\r\n\r\n", textStart);
            if (bodyStart == -1) bodyStart = emlContent.indexOf("\n\n", textStart);
            if (bodyStart == -1) return null;

            bodyStart += (emlContent.charAt(bodyStart + 1) == '\r') ? 4 : 2;
            int bodyEnd = findPartEnd(emlContent, bodyStart);

            return emlContent.substring(bodyStart, bodyEnd).trim();

        } catch (Exception e) {
            return null;
        }
    }

    private static int findPartEnd(String content, int start) {
        String[] lines = content.substring(start).split("\r?\n");
        StringBuilder result = new StringBuilder();

        for (String line : lines) {
            if (line.startsWith("--") && line.length() > 10) break;
            result.append(line).append("\n");
        }

        return start + result.length();
    }

    private static String convertTextToHtml(String textBody) {
        if (textBody == null) return "";

        String html = escapeHtml(textBody);
        html = html.replace("\r\n", "\n").replace("\r", "\n");
        html = html.replace("\n", "<br>\n");

        html =
                html.replaceAll(
                        "(https?://[\\w\\-._~:/?#\\[\\]@!$&'()*+,;=%]+)",
                        "<a href=\"$1\" style=\"color: #1a73e8; text-decoration: underline;\">$1</a>");

        html =
                html.replaceAll(
                        "([a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\\.[a-zA-Z]{2,63})",
                        "<a href=\"mailto:$1\" style=\"color: #1a73e8; text-decoration: underline;\">$1</a>");

        return html;
    }

    private static String processEmailHtmlBody(String htmlBody, EmlToPdfRequest request) {
        if (htmlBody == null) return "";

        String processed = htmlBody;
        boolean modestFormatting = (request != null) && request.isUseModestFormatting();

        // Remove problematic CSS - Corrected regex for Java
        processed = processed.replaceAll("(?i)\\s*position\\s*:\\s*fixed[^;]*;?", "");
        processed = processed.replaceAll("(?i)\\s*position\\s*:\\s*absolute[^;]*;?", "");

        // Potentially remove script and style tags if modestFormatting is true for a cleaner output
        if (modestFormatting) {
            processed = processed.replaceAll("(?i)<script[^>]*>.*?</script>", "");
            processed = processed.replaceAll("(?i)<style[^>]*>.*?</style>", "");
        }

        return processed;
    }

    private static void appendEnhancedStyles(StringBuilder html, EmlToPdfRequest request) {
        int fontSize = 12; // Default font size
        boolean modestFormatting = false;

        if (request != null) {
            modestFormatting = request.isUseModestFormatting();
        }

        String textColor = modestFormatting ? "#000000" : "#202124";
        String backgroundColor = "#ffffff";
        String borderColor = modestFormatting ? "#dddddd" : "#e8eaed";

        html.append("body {\n");
        if (modestFormatting) {
            html.append("  font-family: Arial, sans-serif;\n");
        } else {
            html.append("  font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;\n");
        }
        html.append("  font-size: ").append(fontSize).append("px;\n");
        html.append("  line-height: 1.4;\n");
        html.append("  color: ").append(textColor).append(";\n");
        html.append("  margin: 0;\n");
        html.append("  padding: 16px;\n");
        html.append("  background-color: ").append(backgroundColor).append(";\n");
        html.append("}\n\n");

        html.append(".email-container {\n");
        html.append("  width: 100%;\n");
        html.append("  max-width: 100%;\n");
        html.append("  margin: 0 auto;\n");
        html.append("}\n\n");

        if (!modestFormatting) {
            html.append(".email-header {\n");
            html.append("  padding: 16px 0;\n");
            html.append("  border-bottom: 1px solid ").append(borderColor).append(";\n");
            html.append("  margin-bottom: 16px;\n");
            html.append("}\n\n");

            html.append(".email-header h1 {\n");
            html.append("  margin: 0 0 16px 0;\n");
            html.append("  font-size: ").append(fontSize + 8).append("px;\n");
            html.append("  font-weight: 600;\n");
            html.append("}\n\n");

            html.append(".email-meta div {\n");
            html.append("  margin-bottom: 4px;\n");
            html.append("}\n\n");
        } else {
            html.append(".email-header {\n");
            html.append("  padding-bottom: 10px;\n");
            html.append("  border-bottom: 1px solid ").append(borderColor).append(";\n");
            html.append("  margin-bottom: 10px;\n");
            html.append("}\n\n");
            html.append(".email-header h1 {\n");
            html.append("  margin: 0 0 10px 0;\n");
            html.append("  font-size: ").append(fontSize + 4).append("px;\n");
            html.append("  font-weight: bold;\n");
            html.append("}\n\n");
            html.append(".email-meta div {\n");
            html.append("  margin-bottom: 2px;\n");
            html.append("  font-size: ").append(fontSize - 1).append("px;\n");
            html.append("}\n\n");
        }

        html.append(".email-body {\n");
        html.append("  word-wrap: break-word;\n");
        html.append("}\n\n");

        if (!modestFormatting) {
            html.append(".advanced-features-notice {\n");
            html.append("  margin-top: 20px;\n");
            html.append("  padding: 10px;\n");
            html.append("  background-color: #f0f8ff;\n");
            html.append("  border: 1px solid #cce7ff;\n");
            html.append("  border-radius: 4px;\n");
            html.append("  font-style: italic;\n");
            html.append("  color: #0066cc;\n");
            html.append("}\n\n");

            html.append(".attachment-section {\n");
            html.append("  margin-top: 20px;\n");
            html.append("  padding: 15px;\n");
            html.append("  background-color: #f8f9fa;\n");
            html.append("  border: 1px solid #dee2e6;\n");
            html.append("  border-radius: 6px;\n");
            html.append("}\n\n");

            html.append(".attachment-section h3 {\n");
            html.append("  margin: 0 0 10px 0;\n");
            html.append("  color: #495057;\n");
            html.append("  font-size: ").append(fontSize + 2).append("px;\n");
            html.append("}\n\n");

            html.append(".attachment-item {\n");
            html.append("  display: flex;\n");
            html.append("  align-items: center;\n");
            html.append("  padding: 8px 0;\n");
            html.append("  border-bottom: 1px solid #e9ecef;\n");
            html.append("}\n\n");

            html.append(".attachment-item:last-child {\n");
            html.append("  border-bottom: none;\n");
            html.append("}\n\n");

            html.append(".attachment-icon {\n");
            html.append("  margin-right: 8px;\n");
            html.append("  font-size: 18px;\n");
            html.append("}\n\n");

            html.append(".attachment-name {\n");
            html.append("  font-weight: 500;\n");
            html.append("  color: #495057;\n");
            html.append("}\n\n");

            html.append(".attachment-type {\n");
            html.append("  color: #6c757d;\n");
            html.append("  font-size: ").append(fontSize - 1).append("px;\n");
            html.append("}\n\n");

            html.append(".attachment-details {\n");
            html.append("  color: #6c757d;\n");
            html.append("  font-size: ").append(fontSize - 1).append("px;\n");
            html.append("  margin-left: 8px;\n");
            html.append("}\n\n");

            html.append(".attachment-embedded {\n");
            html.append("  color: #28a745;\n");
            html.append("  font-size: ").append(fontSize - 2).append("px;\n");
            html.append("  font-weight: bold;\n");
            html.append("  margin-left: 8px;\n");
            html.append("}\n\n");

            html.append(".attachment-inclusion-note {\n");
            html.append("  margin-top: 10px;\n");
            html.append("  padding: 8px;\n");
            html.append("  background-color: #d4edda;\n");
            html.append("  border: 1px solid #c3e6cb;\n");
            html.append("  border-radius: 4px;\n");
            html.append("  color: #155724;\n");
            html.append("}\n\n");

            html.append(".attachment-info-note {\n");
            html.append("  margin-top: 10px;\n");
            html.append("  padding: 8px;\n");
            html.append("  background-color: #fff3cd;\n");
            html.append("  border: 1px solid #ffeaa7;\n");
            html.append("  border-radius: 4px;\n");
            html.append("  color: #856404;\n");
            html.append("}\n\n");
        } else {
            html.append(".attachment-section {\n");
            html.append("  margin-top: 15px;\n");
            html.append("  padding: 10px;\n");
            html.append("  background-color: #f9f9f9;\n");
            html.append("  border: 1px solid #eeeeee;\n");
            html.append("  border-radius: 3px;\n");
            html.append("}\n\n");
            html.append(".attachment-section h3 {\n");
            html.append("  margin: 0 0 8px 0;\n");
            html.append("  font-size: ").append(fontSize + 1).append("px;\n");
            html.append("}\n\n");
            html.append(".attachment-item {\n");
            html.append("  padding: 5px 0;\n");
            html.append("}\n\n");
            html.append(".attachment-icon {\n");
            html.append("  margin-right: 5px;\n");
            html.append("}\n\n");
            html.append(".attachment-details, .attachment-type {\n");
            html.append("  font-size: ").append(fontSize - 2).append("px;\n");
            html.append("  color: #555555;\n");
            html.append("}\n\n");
            html.append(".attachment-inclusion-note, .attachment-info-note {\n");
            html.append("  margin-top: 8px;\n");
            html.append("  padding: 6px;\n");
            html.append("  font-size: ").append(fontSize - 2).append("px;\n");
            html.append("  border-radius: 3px;\n");
            html.append("}\n\n");
            html.append(".attachment-inclusion-note {\n");
            html.append("  background-color: #e6ffed;\n");
            html.append("  border: 1px solid #d4f7dc;\n");
            html.append("  color: #006420;\n");
            html.append("}\n\n");
            html.append(".attachment-info-note {\n");
            html.append("  background-color: #fff9e6;\n");
            html.append("  border: 1px solid #fff0c2;\n");
            html.append("  color: #664d00;\n");
            html.append("}\n\n");
            html.append(".attachment-link-container {\n");
            html.append("  display: flex;\n");
            html.append("  align-items: center;\n");
            html.append("  padding: 8px;\n");
            html.append("  background-color: #f8f9fa;\n");
            html.append("  border: 1px solid #dee2e6;\n");
            html.append("  border-radius: 4px;\n");
            html.append("  margin: 4px 0;\n");
            html.append("}\n\n");
            html.append(".attachment-link-container:hover {\n");
            html.append("  background-color: #e9ecef;\n");
            html.append("}\n\n");
            html.append(".attachment-note {\n");
            html.append("  font-size: ").append(fontSize - 3).append("px;\n");
            html.append("  color: #6c757d;\n");
            html.append("  font-style: italic;\n");
            html.append("  margin-left: 8px;\n");
            html.append("}\n\n");
        }

        // Basic image styling: ensure images are responsive but not overly constrained.
        html.append("img {\n");
        html.append("  max-width: 100%;\n"); // Make images responsive to container width
        html.append("  height: auto;\n"); // Maintain aspect ratio
        html.append("  display: block;\n"); // Avoid extra space below images
        html.append("}\n\n");
    }

    private static String escapeHtml(String text) {
        if (text == null) return "";
        return text.replace("&", "&amp;")
                .replace("<", "&lt;")
                .replace(">", "&gt;")
                .replace("\"", "&quot;")
                .replace("'", "&#39;");
    }

    private static stirling.software.common.model.api.converters.HTMLToPdfRequest createHtmlRequest(
            EmlToPdfRequest request) {
        stirling.software.common.model.api.converters.HTMLToPdfRequest htmlRequest =
                new stirling.software.common.model.api.converters.HTMLToPdfRequest();
        if (request != null) {
            htmlRequest.setFileInput(request.getFileInput());

            // Calculate zoom based on font size if adaptive font sizing is enabled
            float zoom = 1.0f;
            htmlRequest.setZoom(zoom);
        } else {
            htmlRequest.setZoom(1.0f);
        }
        return htmlRequest;
    }

    private static EmailContent extractEmailContentAdvanced(
            Object message, EmlToPdfRequest request) {
        EmailContent content = new EmailContent();

        try {
            Class<?> messageClass = message.getClass();

            // Extract headers via reflection
            java.lang.reflect.Method getSubject = messageClass.getMethod("getSubject");
            content.subject = (String) getSubject.invoke(message);
            if (content.subject == null) content.subject = "No Subject";

            java.lang.reflect.Method getFrom = messageClass.getMethod("getFrom");
            Object[] fromAddresses = (Object[]) getFrom.invoke(message);
            content.from =
                    fromAddresses != null && fromAddresses.length > 0
                            ? fromAddresses[0].toString()
                            : "";

            java.lang.reflect.Method getAllRecipients = messageClass.getMethod("getAllRecipients");
            Object[] recipients = (Object[]) getAllRecipients.invoke(message);
            content.to =
                    recipients != null && recipients.length > 0 ? recipients[0].toString() : "";

            java.lang.reflect.Method getSentDate = messageClass.getMethod("getSentDate");
            content.date = (Date) getSentDate.invoke(message);

            // Extract content
            java.lang.reflect.Method getContent = messageClass.getMethod("getContent");
            Object messageContent = getContent.invoke(message);

            if (messageContent instanceof String stringContent) {
                java.lang.reflect.Method getContentType = messageClass.getMethod("getContentType");
                String contentType = (String) getContentType.invoke(message);
                if (contentType != null && contentType.toLowerCase().contains("text/html")) {
                    content.htmlBody = stringContent;
                } else {
                    content.textBody = stringContent;
                }
            } else {
                // Handle multipart content
                try {
                    Class<?> multipartClass = Class.forName("jakarta.mail.Multipart");
                    if (multipartClass.isInstance(messageContent)) {
                        processMultipartAdvanced(messageContent, content, request);
                    }
                } catch (Exception e) {
                }
            }

        } catch (Exception e) {
            content.subject = "Email Conversion";
            content.from = "Unknown";
            content.to = "Unknown";
            content.textBody = "Email content could not be parsed with advanced processing";
        }

        return content;
    }

    private static void processMultipartAdvanced(
            Object multipart, EmailContent content, EmlToPdfRequest request) {
        try {
            Class<?> multipartClass = multipart.getClass();
            java.lang.reflect.Method getCount = multipartClass.getMethod("getCount");
            int count = (Integer) getCount.invoke(multipart);

            java.lang.reflect.Method getBodyPart =
                    multipartClass.getMethod("getBodyPart", int.class);

            for (int i = 0; i < count; i++) {
                Object part = getBodyPart.invoke(multipart, i);
                processPartAdvanced(part, content, request);
            }

        } catch (Exception e) {
        }
    }

    private static void processPartAdvanced(
            Object part, EmailContent content, EmlToPdfRequest request) {
        try {
            Class<?> partClass = part.getClass();
            java.lang.reflect.Method isMimeType = partClass.getMethod("isMimeType", String.class);
            java.lang.reflect.Method getContent = partClass.getMethod("getContent");
            java.lang.reflect.Method getDisposition = partClass.getMethod("getDisposition");
            java.lang.reflect.Method getFileName = partClass.getMethod("getFileName");
            java.lang.reflect.Method getContentType = partClass.getMethod("getContentType");
            java.lang.reflect.Method getHeader = partClass.getMethod("getHeader", String.class);

            Object disposition = getDisposition.invoke(part);
            String filename = (String) getFileName.invoke(part);
            String contentType = (String) getContentType.invoke(part);

            if ((Boolean) isMimeType.invoke(part, "text/plain") && disposition == null) {
                content.textBody = (String) getContent.invoke(part);
            } else if ((Boolean) isMimeType.invoke(part, "text/html") && disposition == null) {
                content.htmlBody = (String) getContent.invoke(part);
            } else if ("attachment".equalsIgnoreCase((String) disposition)
                    || (filename != null && !filename.trim().isEmpty())) {

                content.attachmentCount++;

                // Always extract basic attachment metadata for display
                if (filename != null && !filename.trim().isEmpty()) {
                    // Create attachment with metadata only
                    EmailAttachment attachment = new EmailAttachment();
                    attachment.setFilename(filename);
                    attachment.setContentType(contentType);

                    // Check if it's an embedded image
                    String[] contentIdHeaders = (String[]) getHeader.invoke(part, "Content-ID");
                    if (contentIdHeaders != null && contentIdHeaders.length > 0) {
                        attachment.setEmbedded(true);
                    }

                    // Extract attachment data only if attachments should be included
                    if (request != null && request.isIncludeAttachments()) {
                        try {
                            Object attachmentContent = getContent.invoke(part);
                            byte[] attachmentData = null;

                            if (attachmentContent instanceof java.io.InputStream) {
                                try (java.io.InputStream inputStream =
                                        (java.io.InputStream) attachmentContent) {
                                    attachmentData = inputStream.readAllBytes();
                                }
                            } else if (attachmentContent instanceof byte[]) {
                                attachmentData = (byte[]) attachmentContent;
                            } else if (attachmentContent instanceof String) {
                                attachmentData =
                                        ((String) attachmentContent)
                                                .getBytes(StandardCharsets.UTF_8);
                            }

                            if (attachmentData != null) {
                                // Check size limit
                                long maxSizeMB = request.getMaxAttachmentSizeMB();
                                long maxSizeBytes = maxSizeMB * 1024 * 1024;

                                if (attachmentData.length <= maxSizeBytes) {
                                    attachment.setData(attachmentData);
                                    attachment.setSizeBytes(attachmentData.length);
                                } else {
                                    // Still show attachment info even if too large
                                    attachment.setSizeBytes(attachmentData.length);
                                }
                            }
                        } catch (Exception e) {
                        }
                    }

                    // Add attachment to list for display (with or without data)
                    content.attachments.add(attachment);
                }
            } else if ((Boolean) isMimeType.invoke(part, "multipart/*")) {
                // Handle nested multipart content
                try {
                    Object multipartContent = getContent.invoke(part);
                    Class<?> multipartClass = Class.forName("jakarta.mail.Multipart");
                    if (multipartClass.isInstance(multipartContent)) {
                        processMultipartAdvanced(multipartContent, content, request);
                    }
                } catch (Exception e) {
                }
            }

        } catch (Exception e) {
        }
    }

    private static String generateEnhancedEmailHtml(EmailContent content, EmlToPdfRequest request) {
        StringBuilder html = new StringBuilder();

        html.append("<!DOCTYPE html>\n");
        html.append("<html><head><meta charset=\"UTF-8\">\n");
        html.append("<title>").append(escapeHtml(content.subject)).append("</title>\n");
        html.append("<style>\n");
        appendEnhancedStyles(html, request);
        html.append("</style>\n");
        html.append("</head><body>\n");

        html.append("<div class=\"email-container\">\n");
        html.append("<div class=\"email-header\">\n");
        html.append("<h1>").append(escapeHtml(content.subject)).append("</h1>\n");
        html.append("<div class=\"email-meta\">\n");
        html.append("<div><strong>From:</strong> ")
                .append(escapeHtml(content.from))
                .append("</div>\n");
        html.append("<div><strong>To:</strong> ").append(escapeHtml(content.to)).append("</div>\n");

        if (content.date != null) {
            html.append("<div><strong>Date:</strong> ")
                    .append(formatEmailDate(content.date))
                    .append("</div>\n");
        }
        html.append("</div></div>\n");

        html.append("<div class=\"email-body\">\n");
        if (content.htmlBody != null && !content.htmlBody.trim().isEmpty()) {
            html.append(processEmailHtmlBody(content.htmlBody, request));
        } else if (content.textBody != null && !content.textBody.trim().isEmpty()) {
            html.append("<div class=\"text-body\">");
            html.append(convertTextToHtml(content.textBody));
            html.append("</div>");
        } else {
            html.append("<div class=\"no-content\">");
            html.append("<p><em>No content available</em></p>");
            html.append("</div>");
        }
        html.append("</div>\n");

        if (content.attachmentCount > 0 || !content.attachments.isEmpty()) {
            html.append("<div class=\"attachment-section\">\n");
            int displayedAttachmentCount =
                    content.attachmentCount > 0
                            ? content.attachmentCount
                            : content.attachments.size();
            html.append("<h3>Attachments (").append(displayedAttachmentCount).append(")</h3>\n");

            if (!content.attachments.isEmpty()) {
                for (EmailAttachment attachment : content.attachments) {
                    html.append("<div class=\"attachment-item\">\n");

                    // Create paperclip emoji with unique ID for embedded file linking
                    String uniqueId =
                            "attachment_"
                                    + attachment.filename.hashCode()
                                    + "_"
                                    + System.nanoTime();
                    attachment.embeddedFilename =
                            attachment.embeddedFilename != null
                                    ? attachment.embeddedFilename
                                    : attachment.filename;
                    html.append("<span class=\"attachment-link\" id=\"")
                            .append(uniqueId)
                            .append("\" data-filename=\"")
                            .append(escapeHtml(attachment.embeddedFilename))
                            .append("\">\n")
                            .append("<span class=\"attachment-name\">")
                            .append(escapeHtml(attachment.filename))
                            .append("</span>\n");

                    String sizeStr = formatFileSize(attachment.sizeBytes);
                    html.append("<span class=\"attachment-details\"> (").append(sizeStr);
                    if (attachment.contentType != null && !attachment.contentType.isEmpty()) {
                        html.append(", ").append(escapeHtml(attachment.contentType));
                    }
                    html.append(")</span>\n");

                    html.append("</span>\n");
                    html.append("</div>\n");
                }
            }

            if (request.isIncludeAttachments()) {
                html.append("<div class=\"attachment-info-note\">\n");
                html.append(
                        "<p><em>Attachments saved as external files and linked in this PDF. Click links to open externally.</em></p>\n");
                html.append("</div>\n");
            } else {
                html.append("<div class=\"attachment-info-note\">\n");
                html.append(
                        "<p><em>Attachment information displayed - files not included in PDF.</em></p>\n");
                html.append("</div>\n");
            }

            html.append("</div>\n");
        }

        html.append("</div>\n");
        html.append("</body></html>");

        return html.toString();
    }

    private static byte[] attachFilesToPdf(byte[] pdfBytes, List<EmailAttachment> attachments)
            throws IOException {
        try (PDDocument document = Loader.loadPDF(pdfBytes);
                ByteArrayOutputStream outputStream = new ByteArrayOutputStream()) {

            if (attachments == null || attachments.isEmpty()) {
                document.save(outputStream);
                return outputStream.toByteArray();
            }

            List<String> embeddedFiles = new ArrayList<>();

            // Set up the embedded files name tree once
            if (document.getDocumentCatalog().getNames() == null) {
                document.getDocumentCatalog()
                        .setNames(new PDDocumentNameDictionary(document.getDocumentCatalog()));
            }

            PDDocumentNameDictionary names = document.getDocumentCatalog().getNames();
            if (names.getEmbeddedFiles() == null) {
                names.setEmbeddedFiles(new PDEmbeddedFilesNameTreeNode());
            }

            PDEmbeddedFilesNameTreeNode efTree = names.getEmbeddedFiles();
            Map<String, PDComplexFileSpecification> efMap = efTree.getNames();
            if (efMap == null) {
                efMap = new HashMap<>();
            }

            // Embed each attachment directly into the PDF
            for (EmailAttachment attachment : attachments) {
                if (attachment.data == null || attachment.data.length == 0) {
                    continue;
                }

                try {
                    // Generate unique filename
                    String filename = attachment.filename;
                    if (filename == null || filename.trim().isEmpty()) {
                        filename = "attachment_" + System.currentTimeMillis();
                        if (attachment.contentType != null
                                && attachment.contentType.contains("/")) {
                            String[] parts = attachment.contentType.split("/");
                            if (parts.length > 1) {
                                filename += "." + parts[1];
                            }
                        }
                    }

                    // Ensure unique filename
                    String uniqueFilename = filename;
                    int counter = 1;
                    while (embeddedFiles.contains(uniqueFilename)
                            || efMap.containsKey(uniqueFilename)) {
                        String extension = "";
                        String baseName = filename;
                        int lastDot = filename.lastIndexOf('.');
                        if (lastDot > 0) {
                            extension = filename.substring(lastDot);
                            baseName = filename.substring(0, lastDot);
                        }
                        uniqueFilename = baseName + "_" + counter + extension;
                        counter++;
                    }

                    // Create embedded file
                    PDEmbeddedFile embeddedFile =
                            new PDEmbeddedFile(document, new ByteArrayInputStream(attachment.data));
                    embeddedFile.setSize(attachment.data.length);
                    embeddedFile.setCreationDate(new GregorianCalendar());
                    if (attachment.contentType != null) {
                        embeddedFile.setSubtype(attachment.contentType);
                    }

                    // Create file specification
                    PDComplexFileSpecification fileSpec = new PDComplexFileSpecification();
                    fileSpec.setFile(uniqueFilename);
                    fileSpec.setEmbeddedFile(embeddedFile);
                    if (attachment.contentType != null) {
                        fileSpec.setFileDescription("Email attachment: " + uniqueFilename);
                    }

                    // Add to the map (but don't set it yet)
                    efMap.put(uniqueFilename, fileSpec);
                    embeddedFiles.add(uniqueFilename);

                    // Store the filename for annotation creation
                    attachment.embeddedFilename = uniqueFilename;

                } catch (Exception e) {
                    // Log error but continue with other attachments
                    System.err.println(
                            "Failed to embed attachment: "
                                    + attachment.filename
                                    + " - "
                                    + e.getMessage());
                }
            }

            // Set the complete map once at the end
            if (!efMap.isEmpty()) {
                efTree.setNames(efMap);
            }

            // Add attachment annotations to the first page for each embedded file
            if (!embeddedFiles.isEmpty()) {
                addAttachmentAnnotationsToDocument(document, attachments);
            }

            document.save(outputStream);
            return outputStream.toByteArray();
        }
    }

    private static void addAttachmentAnnotationsToDocument(
            PDDocument document, List<EmailAttachment> attachments) throws IOException {
        if (document.getNumberOfPages() == 0 || attachments.isEmpty()) {
            return;
        }

        // Search for attachment text patterns across all pages to place annotations accurately
        for (int pageIndex = 0; pageIndex < document.getNumberOfPages(); pageIndex++) {
            PDPage page = document.getPage(pageIndex);

            try {
                // Use custom text stripper to get precise character positions
                PaperclipPositionStripper stripper = new PaperclipPositionStripper(attachments);
                stripper.setStartPage(pageIndex + 1);
                stripper.setEndPage(pageIndex + 1);
                stripper.getText(document);

                // Check if this page contains attachment filenames and add annotations
                List<PaperclipPosition> positions = stripper.getPaperclipPositions();
                if (!positions.isEmpty()) {
                    addAttachmentAnnotationsToPage(document, page, attachments, positions);
                    break; // Only add annotations to the page that contains the attachments section
                }
            } catch (Exception e) {
                System.err.println(
                        "Failed to process page "
                                + pageIndex
                                + " for attachment annotations: "
                                + e.getMessage());
            }
        }
    }

    private static void addAttachmentAnnotationsToPage(
            PDDocument document,
            PDPage page,
            List<EmailAttachment> attachments,
            List<PaperclipPosition> paperclipPositions)
            throws IOException {

        List<org.apache.pdfbox.pdmodel.interactive.annotation.PDAnnotation> annotations =
                new ArrayList<>(page.getAnnotations());

        PDRectangle pageSize = page.getMediaBox();
        float pageHeight = pageSize.getHeight();

        try {
            // Match paperclip positions with attachments
            int attachmentIndex = 0;
            for (PaperclipPosition position : paperclipPositions) {
                if (attachmentIndex >= attachments.size()) {
                    break;
                }

                EmailAttachment attachment = attachments.get(attachmentIndex);
                if (attachment.embeddedFilename == null) {
                    attachmentIndex++;
                    continue;
                }

                try {
                    // Create file attachment annotation at the exact paperclip position
                    PDAnnotationFileAttachment fileAnnotation = new PDAnnotationFileAttachment();

                    // PDFBox coordinates: (0,0) is bottom-left, Y increases upward
                    // TextPosition.getYDirAdj() gives the baseline of the text
                    // We need to position the annotation slightly above the baseline to cover the
                    // emoji
                    float annotationX = position.x;
                    float annotationY = position.y; // Use the exact Y position from TextPosition
                    float annotationWidth =
                            Math.max(position.width, 16f); // Ensure minimum size for click area
                    float annotationHeight = Math.max(position.height, 16f);

                    // Debug output
                    System.out.println(
                            "Creating annotation for "
                                    + attachment.filename
                                    + " at: x="
                                    + annotationX
                                    + ", y="
                                    + annotationY
                                    + ", width="
                                    + annotationWidth
                                    + ", height="
                                    + annotationHeight);

                    // Set the rectangle for the annotation to cover the paperclip emoji area
                    PDRectangle annotationRect =
                            new PDRectangle(
                                    annotationX, annotationY, annotationWidth, annotationHeight);
                    fileAnnotation.setRectangle(annotationRect);

                    // Set the file specification to point to the embedded file
                    PDComplexFileSpecification fileSpec = new PDComplexFileSpecification();
                    fileSpec.setFile(attachment.embeddedFilename);
                    fileAnnotation.setFile(fileSpec);

                    // Set annotation properties
                    fileAnnotation.setContents("Click to open: " + attachment.filename);
                    fileAnnotation.setAnnotationName("EmbeddedFile_" + attachment.embeddedFilename);

                    // Set paperclip icon
                    fileAnnotation.setSubject(PDAnnotationFileAttachment.ATTACHMENT_NAME_PAPERCLIP);

                    // Add annotation to page
                    annotations.add(fileAnnotation);
                    attachmentIndex++;

                } catch (Exception e) {
                    System.err.println(
                            "Failed to create file attachment annotation for: "
                                    + attachment.filename
                                    + " - "
                                    + e.getMessage());
                    attachmentIndex++;
                }
            }

            // Fallback: if we have more attachments than paperclip positions,
            // place remaining annotations in the attachments area
            if (attachmentIndex < attachments.size()) {
                float fallbackY = pageHeight * 0.4f; // Position in attachments area
                float leftMargin = 72f; // 1 inch from left margin

                for (int i = attachmentIndex; i < attachments.size(); i++) {
                    EmailAttachment attachment = attachments.get(i);
                    if (attachment.embeddedFilename == null) continue;

                    try {
                        PDAnnotationFileAttachment fileAnnotation =
                                new PDAnnotationFileAttachment();

                        // Place annotations in a column within the attachments section
                        PDRectangle annotationRect =
                                new PDRectangle(
                                        leftMargin + 20f, // Indent slightly from margin
                                        Math.max(fallbackY - ((i - attachmentIndex) * 25f), 50f),
                                        14f, // Size to match emoji
                                        14f);
                        fileAnnotation.setRectangle(annotationRect);

                        PDComplexFileSpecification fileSpec = new PDComplexFileSpecification();
                        fileSpec.setFile(attachment.embeddedFilename);
                        fileAnnotation.setFile(fileSpec);

                        fileAnnotation.setContents("Click to open: " + attachment.filename);
                        fileAnnotation.setAnnotationName(
                                "EmbeddedFile_" + attachment.embeddedFilename);
                        fileAnnotation.setSubject(
                                PDAnnotationFileAttachment.ATTACHMENT_NAME_PAPERCLIP);

                        annotations.add(fileAnnotation);

                    } catch (Exception e) {
                        System.err.println(
                                "Failed to create fallback annotation for: " + attachment.filename);
                    }
                }
            }

        } catch (Exception e) {
            System.err.println(
                    "Failed to process paperclip positions for annotations: " + e.getMessage());
        }

        page.setAnnotations(annotations);
    }

    private static String formatEmailDate(Date date) {
        if (date == null) return "";
        java.text.SimpleDateFormat formatter =
                new java.text.SimpleDateFormat("EEE, MMM d, yyyy 'at' h:mm a", Locale.ENGLISH);
        return formatter.format(date);
    }

    private static String formatFileSize(long bytes) {
        if (bytes < 1024) return bytes + " B";
        int exp = (int) (Math.log(bytes) / Math.log(1024));
        String pre = "KMGTPE".charAt(exp - 1) + "";
        return String.format("%.1f %sB", bytes / Math.pow(1024, exp), pre);
    }

    @Data
    @NoArgsConstructor
    @AllArgsConstructor
    private static class PaperclipPosition {
        private float x;
        private float y;
        private float width;
        private float height;
        private String associatedFilename;
    }

    /** Custom PDFTextStripper that captures the exact positions of attachment filenames */
    private static class PaperclipPositionStripper extends PDFTextStripper {
        private List<PaperclipPosition> paperclipPositions = new ArrayList<>();
        private List<EmailAttachment> attachments;

        public PaperclipPositionStripper(List<EmailAttachment> attachments) throws IOException {
            super();
            this.attachments = attachments;
        }

        @Override
        protected void writeString(String string, List<TextPosition> textPositions)
                throws IOException {
            super.writeString(string, textPositions);

            // Look for attachment filenames in the text
            if (attachments != null) {
                for (EmailAttachment attachment : attachments) {
                    String filename = attachment.filename;
                    if (filename != null && !filename.isEmpty()) {
                        int startIndex = string.indexOf(filename);
                        if (startIndex >= 0 && startIndex < textPositions.size()) {
                            // Found an attachment filename, capture its position
                            TextPosition textPos = textPositions.get(startIndex);
                            PaperclipPosition position = new PaperclipPosition();

                            // Use the position where the filename starts
                            position.x = textPos.getXDirAdj();
                            position.y = textPos.getYDirAdj() - textPos.getHeightDir() / 2;

                            // Calculate width based on filename length
                            float totalWidth = 0;
                            int endIndex =
                                    Math.min(startIndex + filename.length(), textPositions.size());
                            for (int i = startIndex; i < endIndex; i++) {
                                totalWidth += textPositions.get(i).getWidthDirAdj();
                            }

                            position.width = Math.max(totalWidth, 16f);
                            position.height = Math.max(textPos.getHeightDir(), 16f);
                            position.associatedFilename = filename;

                            paperclipPositions.add(position);

                            // Debug output
                            System.out.println(
                                    "Found attachment filename '"
                                            + filename
                                            + "' at: x="
                                            + position.x
                                            + ", y="
                                            + position.y
                                            + ", width="
                                            + position.width
                                            + ", height="
                                            + position.height);
                        }
                    }
                }
            }
        }

        public List<PaperclipPosition> getPaperclipPositions() {
            return paperclipPositions;
        }
    }

    @Data
    @NoArgsConstructor
    private static class EmailContent {
        private String subject = "";
        private String from = "";
        private String to = "";
        private Date date;
        private String textBody = "";
        private String htmlBody = "";
        private int attachmentCount = 0;
        private List<EmailAttachment> attachments = new ArrayList<>();
    }

    @Data
    @NoArgsConstructor
    @AllArgsConstructor
    private static class EmailAttachment {
        private String filename = "";
        private String contentType = "";
        private byte[] data;
        private long sizeBytes;
        private boolean embedded = false;
        private String externalFilePath;
        private String embeddedFilename;
    }
}
