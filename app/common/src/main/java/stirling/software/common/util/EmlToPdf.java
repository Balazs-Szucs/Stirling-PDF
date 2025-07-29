package stirling.software.common.util;

import static stirling.software.common.util.AttachmentUtils.setCatalogViewerPreferences;

import java.io.*;
import java.lang.reflect.Constructor;
import java.lang.reflect.Method;
import java.nio.charset.Charset;
import java.nio.charset.StandardCharsets;
import java.util.ArrayList;
import java.util.Base64;
import java.util.Date;
import java.util.GregorianCalendar;
import java.util.HashMap;
import java.util.List;
import java.util.Locale;
import java.util.Map;
import java.util.Properties;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.pdmodel.PDDocumentCatalog;
import org.apache.pdfbox.pdmodel.PDDocumentNameDictionary;
import org.apache.pdfbox.pdmodel.PDEmbeddedFilesNameTreeNode;
import org.apache.pdfbox.pdmodel.PDPage;
import org.apache.pdfbox.pdmodel.PageMode;
import org.apache.pdfbox.pdmodel.common.PDRectangle;
import org.apache.pdfbox.pdmodel.common.filespecification.PDComplexFileSpecification;
import org.apache.pdfbox.pdmodel.common.filespecification.PDEmbeddedFile;
import org.apache.pdfbox.pdmodel.interactive.annotation.PDAnnotationFileAttachment;
import org.apache.pdfbox.pdmodel.interactive.annotation.PDAppearanceDictionary;
import org.apache.pdfbox.pdmodel.interactive.annotation.PDAppearanceStream;
import org.jetbrains.annotations.NotNull;
import org.jetbrains.annotations.Nullable;
import org.springframework.web.multipart.MultipartFile;

import lombok.Data;
import lombok.Getter;
import lombok.experimental.UtilityClass;
import lombok.extern.slf4j.Slf4j;

import stirling.software.common.model.api.converters.EmlToPdfRequest;
import stirling.software.common.model.api.converters.HTMLToPdfRequest;
import stirling.software.common.service.CustomPDFDocumentFactory;

@Slf4j
@UtilityClass
public class EmlToPdf {

    private static final class StyleConstants {
        // Font and layout constants
        static final int DEFAULT_FONT_SIZE = 12;
        static final String DEFAULT_FONT_FAMILY = "Helvetica, sans-serif";
        static final float DEFAULT_LINE_HEIGHT = 1.4f;
        static final String DEFAULT_ZOOM = "1.0";

        // Color constants - aligned with application theme
        static final String DEFAULT_TEXT_COLOR = "#202124";
        static final String DEFAULT_BACKGROUND_COLOR = "#ffffff";
        static final String DEFAULT_BORDER_COLOR = "#e8eaed";
        static final String ATTACHMENT_BACKGROUND_COLOR = "#f9f9f9";
        static final String ATTACHMENT_BORDER_COLOR = "#eeeeee";

        // Size constants for PDF annotations
        static final float ATTACHMENT_ICON_WIDTH = 12f;
        static final float ATTACHMENT_ICON_HEIGHT = 14f;
        static final float ANNOTATION_X_OFFSET = 2f;
        static final float ANNOTATION_Y_OFFSET = 10f;

        // Content validation constants
        static final int EML_CHECK_LENGTH = 8192;
        static final int MIN_HEADER_COUNT_FOR_VALID_EML = 2;

        private StyleConstants() {}
    }

    private static final class MimeConstants {
        static final Pattern MIME_ENCODED_PATTERN =
                Pattern.compile("=\\?([^?]+)\\?([BbQq])\\?([^?]*)\\?=");
        static final String ATTACHMENT_MARKER = "@";

        // Jakarta Mail disposition constants (when available via reflection)
        static final String DISPOSITION_ATTACHMENT = "attachment";

        // Common MIME types
        static final String TEXT_PLAIN = "text/plain";
        static final String TEXT_HTML = "text/html";
        static final String MULTIPART_PREFIX = "multipart/";

        // Common headers
        static final String HEADER_CONTENT_TYPE = "content-type:";
        static final String HEADER_CONTENT_DISPOSITION = "content-disposition:";
        static final String HEADER_CONTENT_TRANSFER_ENCODING = "content-transfer-encoding:";
        static final String HEADER_CONTENT_ID = "Content-ID";
        static final String HEADER_SUBJECT = "Subject:";
        static final String HEADER_FROM = "From:";
        static final String HEADER_TO = "To:";
        static final String HEADER_CC = "Cc:";
        static final String HEADER_BCC = "Bcc:";
        static final String HEADER_DATE = "Date:";

        private MimeConstants() {}
    }

    private static final class MimeTypeDetector {
        private static final Map<String, String> EXTENSION_TO_MIME_TYPE =
                Map.of(
                        ".png", "image/png",
                        ".jpg", "image/jpeg",
                        ".jpeg", "image/jpeg",
                        ".gif", "image/gif",
                        ".bmp", "image/bmp",
                        ".webp", "image/webp",
                        ".svg", "image/svg+xml",
                        ".ico", "image/x-icon",
                        ".tiff", "image/tiff",
                        ".tif", "image/tiff");

        static String detectMimeType(String filename, String existingMimeType) {
            if (existingMimeType != null && !existingMimeType.isEmpty()) {
                return existingMimeType;
            }

            if (filename != null) {
                String lowerFilename = filename.toLowerCase();
                for (Map.Entry<String, String> entry : EXTENSION_TO_MIME_TYPE.entrySet()) {
                    if (lowerFilename.endsWith(entry.getKey())) {
                        return entry.getValue();
                    }
                }
            }

            return "image/png"; // Safe fallback
        }

        private MimeTypeDetector() {}
    }

    // Cached Jakarta Mail availability check and reflection methods
    private static Boolean jakartaMailAvailable = null;
    private static Method mimeUtilityDecodeTextMethod = null;
    private static boolean mimeUtilityChecked = false;

    private static boolean isJakartaMailAvailable() {
        if (jakartaMailAvailable == null) {
            try {
                // Check for core Jakarta Mail classes
                Class.forName("jakarta.mail.internet.MimeMessage");
                Class.forName("jakarta.mail.Session");
                Class.forName("jakarta.mail.internet.MimeUtility");
                Class.forName("jakarta.mail.internet.MimePart");
                Class.forName("jakarta.mail.internet.MimeMultipart");
                Class.forName("jakarta.mail.Multipart");
                Class.forName("jakarta.mail.Part");

                jakartaMailAvailable = true;
                log.debug("Jakarta Mail libraries are available");
            } catch (ClassNotFoundException e) {
                jakartaMailAvailable = false;
                log.debug("Jakarta Mail libraries are not available, using basic parsing");
            }
        }
        return jakartaMailAvailable;
    }

    public static String convertEmlToHtml(byte[] emlBytes, EmlToPdfRequest request)
            throws IOException {
        validateEmlInput(emlBytes);

        if (isJakartaMailAvailable()) {
            return convertEmlToHtmlAdvanced(emlBytes, request);
        } else {
            return convertEmlToHtmlBasic(emlBytes, request, null);
        }
    }

    public static byte[] convertEmlToPdf(
            String weasyprintPath,
            EmlToPdfRequest request,
            byte[] emlBytes,
            String fileName,
            stirling.software.common.service.CustomPDFDocumentFactory pdfDocumentFactory,
            TempFileManager tempFileManager,
            CustomHtmlSanitizer customHtmlSanitizer)
            throws IOException, InterruptedException {

        validateEmlInput(emlBytes);

        try {
            // Generate HTML representation
            EmailContent emailContent = null;
            String htmlContent;

            if (isJakartaMailAvailable()) {
                emailContent = extractEmailContentAdvanced(emlBytes, request, customHtmlSanitizer);
                htmlContent = generateEnhancedEmailHtml(emailContent, request, customHtmlSanitizer);
            } else {
                htmlContent = convertEmlToHtmlBasic(emlBytes, request, customHtmlSanitizer);
            }

            // Convert HTML to PDF
            byte[] pdfBytes =
                    convertHtmlToPdf(
                            weasyprintPath,
                            request,
                            htmlContent,
                            tempFileManager,
                            customHtmlSanitizer);

            // Attach files if available and requested
            if (shouldAttachFiles(emailContent, request)) {
                pdfBytes =
                        attachFilesToPdf(
                                pdfBytes, emailContent.getAttachments(), pdfDocumentFactory);
            }

            return pdfBytes;

        } catch (IOException | InterruptedException e) {
            log.error("Failed to convert EML to PDF for file: {}", fileName, e);
            throw e;
        } catch (Exception e) {
            log.error("Unexpected error during EML to PDF conversion for file: {}", fileName, e);
            throw new IOException("Conversion failed: " + e.getMessage(), e);
        }
    }

    private static void validateEmlInput(byte[] emlBytes) {
        if (emlBytes == null || emlBytes.length == 0) {
            throw new IllegalArgumentException("EML file is empty or null");
        }

        if (isInvalidEmlFormat(emlBytes)) {
            throw new IllegalArgumentException("Invalid EML file format");
        }
    }

    private static boolean shouldAttachFiles(EmailContent emailContent, EmlToPdfRequest request) {
        return emailContent != null
                && request != null
                && request.isIncludeAttachments()
                && !emailContent.getAttachments().isEmpty();
    }

    private static byte[] convertHtmlToPdf(
            String weasyprintPath,
            EmlToPdfRequest request,
            String htmlContent,
            TempFileManager tempFileManager,
            CustomHtmlSanitizer customHtmlSanitizer)
            throws IOException, InterruptedException {

        HTMLToPdfRequest htmlRequest = createHtmlRequest(request);

        try {
            return FileToPdf.convertHtmlToPdf(
                    weasyprintPath,
                    htmlRequest,
                    htmlContent.getBytes(StandardCharsets.UTF_8),
                    "email.html",
                    tempFileManager,
                    customHtmlSanitizer);
        } catch (IOException | InterruptedException e) {
            log.warn("Initial HTML to PDF conversion failed, trying with simplified HTML");
            String simplifiedHtml = simplifyHtmlContent(htmlContent);
            return FileToPdf.convertHtmlToPdf(
                    weasyprintPath,
                    htmlRequest,
                    simplifiedHtml.getBytes(StandardCharsets.UTF_8),
                    "email.html",
                    tempFileManager,
                    customHtmlSanitizer);
        }
    }

    private static String simplifyHtmlContent(String htmlContent) {
        String simplified = htmlContent.replaceAll("(?i)<script[^>]*>.*?</script>", "");
        simplified = simplified.replaceAll("(?i)<style[^>]*>.*?</style>", "");
        return simplified;
    }

    private static String generateUniqueAttachmentId(String filename) {
        // Simple unique ID generation combining filename hash and timestamp
        return "attachment_"
                + (filename != null ? filename.hashCode() : "unknown")
                + "_"
                + System.nanoTime();
    }

    private static String convertEmlToHtmlBasic(
            byte[] emlBytes, EmlToPdfRequest request, CustomHtmlSanitizer customHtmlSanitizer) {
        if (emlBytes == null || emlBytes.length == 0) {
            throw new IllegalArgumentException("EML file is empty or null");
        }

        String emlContent = new String(emlBytes, StandardCharsets.UTF_8);

        // Basic email parsing
        String subject = extractBasicHeader(emlContent, MimeConstants.HEADER_SUBJECT);
        String from = extractBasicHeader(emlContent, MimeConstants.HEADER_FROM);
        String to = extractBasicHeader(emlContent, MimeConstants.HEADER_TO);
        String cc = extractBasicHeader(emlContent, MimeConstants.HEADER_CC);
        String bcc = extractBasicHeader(emlContent, MimeConstants.HEADER_BCC);
        String date = extractBasicHeader(emlContent, MimeConstants.HEADER_DATE);

        // Try to extract HTML content
        String htmlBody = extractHtmlBody(emlContent);
        if (htmlBody == null) {
            String textBody = extractTextBody(emlContent);
            htmlBody =
                    convertTextToHtml(
                            textBody != null ? textBody : "Email content could not be parsed",
                            customHtmlSanitizer);
        }

        // Generate HTML with custom styling based on request
        StringBuilder html = new StringBuilder();

        // HTML header
        html.append(
                String.format(
                        """
            <!DOCTYPE html>
            <html><head><meta charset="UTF-8">
            <title>%s</title>
            <style>
            """,
                        sanitizeText(subject, customHtmlSanitizer)));

        appendEnhancedStyles(html);

        html.append(
                """
            </style>
            </head><body>
            <div class="email-container">
            <div class="email-header">
            """);

        // Email header content
        html.append(
                String.format(
                        """
            <h1>%s</h1>
            <div class="email-meta">
            <div><strong>From:</strong> %s</div>
            <div><strong>To:</strong> %s</div>
            """,
                        sanitizeText(subject, customHtmlSanitizer),
                        sanitizeText(from, customHtmlSanitizer),
                        sanitizeText(to, customHtmlSanitizer)));

        // Include CC and BCC if present and requested
        if (request != null && request.isIncludeAllRecipients()) {
            if (!cc.trim().isEmpty()) {
                html.append(
                        String.format(
                                "<div><strong>CC:</strong> %s</div>\n",
                                sanitizeText(cc, customHtmlSanitizer)));
            }
            if (!bcc.trim().isEmpty()) {
                html.append(
                        String.format(
                                "<div><strong>BCC:</strong> %s</div>\n",
                                sanitizeText(bcc, customHtmlSanitizer)));
            }
        }

        if (!date.trim().isEmpty()) {
            html.append(
                    String.format(
                            "<div><strong>Date:</strong> %s</div>\n",
                            sanitizeText(date, customHtmlSanitizer)));
        }

        html.append(
                """
            </div></div>
            <div class="email-body">
            """);
        html.append(processEmailHtmlBody(htmlBody, customHtmlSanitizer));
        html.append("</div>\n");

        // Add attachment information - always check for and display attachments
        String attachmentInfo = extractAttachmentInfo(emlContent);
        if (!attachmentInfo.isEmpty()) {
            html.append(
                    """
                <div class="attachment-section">
                <h3>Attachments</h3>
                """);
            html.append(attachmentInfo);

            // Add a status message about attachment inclusion
            if (request != null && request.isIncludeAttachments()) {
                html.append(
                        """
                    <div class="attachment-inclusion-note">
                    <p><strong>Note:</strong> Attachments are saved as external files and linked in this PDF. Click the links to open files externally.</p>
                    </div>
                    """);
            } else {
                html.append(
                        """
                    <div class="attachment-info-note">
                    <p><em>Attachment information displayed - files not included in PDF. Enable 'Include attachments' to embed files.</em></p>
                    </div>
                    """);
            }

            html.append("</div>\n");
        }

        // Show advanced features status if requested
        if (request != null && request.getFileInput().isEmpty()) {
            html.append(
                    """
                <div class="advanced-features-notice">
                <p><em>Note: Some advanced features require Jakarta Mail dependencies.</em></p>
                </div>
                """);
        }

        html.append(
                """
            </div>
            </body></html>
            """);

        return html.toString();
    }

    private static EmailContent extractEmailContentAdvanced(
            byte[] emlBytes, EmlToPdfRequest request) {
        try {
            // Use Jakarta Mail for processing
            Class<?> sessionClass = Class.forName("jakarta.mail.Session");
            Class<?> mimeMessageClass = Class.forName("jakarta.mail.internet.MimeMessage");

            Method getDefaultInstance =
                    sessionClass.getMethod("getDefaultInstance", Properties.class);
            Object session = getDefaultInstance.invoke(null, new Properties());

            // Cast the session object to the proper type for the constructor
            Class<?>[] constructorArgs = new Class<?>[] {sessionClass, InputStream.class};
            Constructor<?> mimeMessageConstructor =
                    mimeMessageClass.getConstructor(constructorArgs);
            Object message =
                    mimeMessageConstructor.newInstance(session, new ByteArrayInputStream(emlBytes));

            return extractEmailContentAdvanced(message, request, null);

        } catch (ReflectiveOperationException e) {
            // Create basic EmailContent from basic processing
            EmailContent content = new EmailContent();
            content.setHtmlBody(convertEmlToHtmlBasic(emlBytes, request, null));
            return content;
        }
    }

    private static EmailContent extractEmailContentAdvanced(
            byte[] emlBytes, EmlToPdfRequest request, CustomHtmlSanitizer customHtmlSanitizer) {
        try {
            // Use Jakarta Mail for processing
            Class<?> sessionClass = Class.forName("jakarta.mail.Session");
            Class<?> mimeMessageClass = Class.forName("jakarta.mail.internet.MimeMessage");

            Method getDefaultInstance =
                    sessionClass.getMethod("getDefaultInstance", Properties.class);
            Object session = getDefaultInstance.invoke(null, new Properties());

            // Cast the session object to the proper type for the constructor
            Class<?>[] constructorArgs = new Class<?>[] {sessionClass, InputStream.class};
            Constructor<?> mimeMessageConstructor =
                    mimeMessageClass.getConstructor(constructorArgs);
            Object message =
                    mimeMessageConstructor.newInstance(session, new ByteArrayInputStream(emlBytes));

            return extractEmailContentAdvanced(message, request, customHtmlSanitizer);

        } catch (ReflectiveOperationException e) {
            // Create basic EmailContent from basic processing
            EmailContent content = new EmailContent();
            content.setHtmlBody(convertEmlToHtmlBasic(emlBytes, request, customHtmlSanitizer));
            return content;
        }
    }

    private static String convertEmlToHtmlAdvanced(byte[] emlBytes, EmlToPdfRequest request) {
        EmailContent content = extractEmailContentAdvanced(emlBytes, request);
        return generateEnhancedEmailHtml(content, request, null);
    }

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
                if (lowerLine.startsWith(MimeConstants.HEADER_CONTENT_TYPE)
                        && lowerLine.contains(MimeConstants.MULTIPART_PREFIX)) {
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
                if (lowerLine.startsWith(MimeConstants.HEADER_CONTENT_TYPE)) {
                    currentContentType =
                            line.substring(MimeConstants.HEADER_CONTENT_TYPE.length()).trim();
                } else if (lowerLine.startsWith(MimeConstants.HEADER_CONTENT_DISPOSITION)) {
                    currentDisposition =
                            line.substring(MimeConstants.HEADER_CONTENT_DISPOSITION.length())
                                    .trim();
                    // Extract filename if present
                    currentFilename = extractFilenameFromDisposition(currentDisposition);
                } else if (lowerLine.startsWith(MimeConstants.HEADER_CONTENT_TRANSFER_ENCODING)) {
                    currentEncoding =
                            line.substring(MimeConstants.HEADER_CONTENT_TRANSFER_ENCODING.length())
                                    .trim();
                } else if (line.startsWith(" ") || line.startsWith("\t")) {
                    // Continuation of previous header
                    if (currentDisposition.contains("filename=")) {
                        currentDisposition += " " + line.trim();
                        currentFilename = extractFilenameFromDisposition(currentDisposition);
                    } else if (!currentContentType.isEmpty()) {
                        currentContentType += " " + line.trim();
                    }
                }
            }

            if (isAttachment(currentDisposition, currentFilename, currentContentType)) {
                addAttachmentToInfo(
                        attachmentInfo, currentFilename, currentContentType, currentEncoding);
            }

        } catch (RuntimeException e) {
            log.warn("Error extracting attachment info: {}", e.getMessage());
        }
        return attachmentInfo.toString();
    }

    private static boolean isAttachment(String disposition, String filename, String contentType) {
        return (disposition.toLowerCase().contains(MimeConstants.DISPOSITION_ATTACHMENT)
                        && !filename.isEmpty())
                || (!filename.isEmpty() && !contentType.toLowerCase().startsWith("text/"))
                || (contentType.toLowerCase().contains("application/") && !filename.isEmpty());
    }

    private static String extractFilenameFromDisposition(String disposition) {
        if (disposition.contains("filename=")) {
            int filenameStart = disposition.toLowerCase().indexOf("filename=") + 9;
            int filenameEnd = disposition.indexOf(";", filenameStart);
            if (filenameEnd == -1) filenameEnd = disposition.length();
            String filename = disposition.substring(filenameStart, filenameEnd).trim();
            filename = filename.replaceAll("^\"|\"$", "");
            // Apply MIME decoding to handle encoded filenames
            return safeMimeDecode(filename);
        }
        return "";
    }

    private static void addAttachmentToInfo(
            StringBuilder attachmentInfo, String filename, String contentType, String encoding) {
        // Create attachment info with paperclip emoji before filename
        attachmentInfo
                .append("<div class=\"attachment-item\">")
                .append("<span class=\"attachment-icon\">")
                .append(MimeConstants.ATTACHMENT_MARKER)
                .append("</span> ")
                .append("<span class=\"attachment-name\">")
                .append(escapeHtml(filename))
                .append("</span>");

        // Add content type and encoding info
        if (!contentType.isEmpty() || !encoding.isEmpty()) {
            attachmentInfo.append(" <span class=\"attachment-details\">(");
            if (!contentType.isEmpty()) {
                attachmentInfo.append(escapeHtml(contentType));
            }
            if (!encoding.isEmpty()) {
                if (!contentType.isEmpty()) attachmentInfo.append(", ");
                attachmentInfo.append("encoding: ").append(escapeHtml(encoding));
            }
            attachmentInfo.append(")</span>");
        }
        attachmentInfo.append("</div>\n");
    }

    private static boolean isInvalidEmlFormat(byte[] emlBytes) {
        try {
            int checkLength = Math.min(emlBytes.length, StyleConstants.EML_CHECK_LENGTH);
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

            return headerCount < StyleConstants.MIN_HEADER_COUNT_FOR_VALID_EML && !hasMimeStructure;

        } catch (RuntimeException e) {
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
                    // Apply MIME header decoding
                    return safeMimeDecode(value.toString());
                }
                if (line.trim().isEmpty()) break;
            }
        } catch (RuntimeException e) {
            log.warn("Error extracting header '{}': {}", headerName, e.getMessage());
        }
        return "";
    }

    private static String extractHtmlBody(String emlContent) {
        try {
            String lowerContent = emlContent.toLowerCase();
            int htmlStart =
                    lowerContent.indexOf(
                            MimeConstants.HEADER_CONTENT_TYPE + " " + MimeConstants.TEXT_HTML);
            if (htmlStart == -1) return null;

            return getString(emlContent, htmlStart);

        } catch (Exception e) {
            return null;
        }
    }

    @Nullable
    private static String getString(String emlContent, int htmlStart) {
        int bodyStart = emlContent.indexOf("\r\n\r\n", htmlStart);
        if (bodyStart == -1) bodyStart = emlContent.indexOf("\n\n", htmlStart);
        if (bodyStart == -1) return null;

        bodyStart += (emlContent.charAt(bodyStart + 1) == '\r') ? 4 : 2;
        int bodyEnd = findPartEnd(emlContent, bodyStart);

        return emlContent.substring(bodyStart, bodyEnd).trim();
    }

    private static String extractTextBody(String emlContent) {
        try {
            String lowerContent = emlContent.toLowerCase();
            int textStart =
                    lowerContent.indexOf(
                            MimeConstants.HEADER_CONTENT_TYPE + " " + MimeConstants.TEXT_PLAIN);
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

            return getString(emlContent, textStart);

        } catch (RuntimeException e) {
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

    private static String convertTextToHtml(
            String textBody, CustomHtmlSanitizer customHtmlSanitizer) {
        if (textBody == null) return "";

        // Use the sanitizer to escape HTML entities properly, or fall back to manual escaping
        String html;
        if (customHtmlSanitizer != null) {
            html = customHtmlSanitizer.sanitize(textBody);
        } else {
            html = escapeHtml(textBody);
        }

        html = html.replace("\r\n", "\n").replace("\r", "\n");
        html = html.replace("\n", "<br>\n");

        // Convert URLs to clickable links
        html =
                html.replaceAll(
                        "(https?://[\\w\\-._~:/?#\\[\\]@!$&'()*+,;=%]+)",
                        "<a href=\"$1\" style=\"color: #1a73e8; text-decoration: underline;\">$1</a>");

        // Convert email addresses to mailto links
        html =
                html.replaceAll(
                        "([a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\\.[a-zA-Z]{2,63})",
                        "<a href=\"mailto:$1\" style=\"color: #1a73e8; text-decoration: underline;\">$1</a>");

        return html;
    }

    private static String processEmailHtmlBody(
            String htmlBody, CustomHtmlSanitizer customHtmlSanitizer) {
        return processEmailHtmlBody(htmlBody, null, customHtmlSanitizer);
    }

    private static String processEmailHtmlBody(
            String htmlBody, EmailContent emailContent, CustomHtmlSanitizer customHtmlSanitizer) {
        if (htmlBody == null) return "";

        // Apply the custom HTML sanitizer to clean the content if available
        String processed;
        if (customHtmlSanitizer != null) {
            processed = customHtmlSanitizer.sanitize(htmlBody);
        } else {
            processed = htmlBody; // No sanitization for public API backward compatibility
        }

        // Remove problematic CSS that might interfere with PDF generation
        processed = processed.replaceAll("(?i)\\s*position\\s*:\\s*fixed[^;]*;?", "");
        processed = processed.replaceAll("(?i)\\s*position\\s*:\\s*absolute[^;]*;?", "");

        if (emailContent != null && !emailContent.getAttachments().isEmpty()) {
            processed = processInlineImages(processed, emailContent);
        }

        return processed;
    }

    private static String processInlineImages(String htmlContent, EmailContent emailContent) {
        if (htmlContent == null || emailContent == null) return htmlContent;

        // Create a map of Content-ID to attachment data
        Map<String, EmailAttachment> contentIdMap = new HashMap<>();
        for (EmailAttachment attachment : emailContent.getAttachments()) {
            if (attachment.isEmbedded()
                    && attachment.getContentId() != null
                    && attachment.getData() != null) {
                contentIdMap.put(attachment.getContentId(), attachment);
            }
        }

        if (contentIdMap.isEmpty()) return htmlContent;

        // Pattern to match cid: references in img src attributes
        Pattern cidPattern =
                Pattern.compile(
                        "(?i)<img[^>]*\\ssrc\\s*=\\s*['\"]cid:([^'\"]+)['\"][^>]*>",
                        Pattern.CASE_INSENSITIVE);
        Matcher matcher = cidPattern.matcher(htmlContent);

        StringBuilder result = new StringBuilder();
        while (matcher.find()) {
            String contentId = matcher.group(1);
            EmailAttachment attachment = contentIdMap.get(contentId);

            if (attachment != null && attachment.getData() != null) {
                // Convert to data URI using improved MIME type detection
                String mimeType =
                        MimeTypeDetector.detectMimeType(
                                attachment.getFilename(), attachment.getContentType());

                String base64Data = Base64.getEncoder().encodeToString(attachment.getData());
                String dataUri = "data:" + mimeType + ";base64," + base64Data;

                // Replace the cid: reference with the data URI
                String replacement =
                        matcher.group(0).replaceFirst("cid:" + Pattern.quote(contentId), dataUri);
                matcher.appendReplacement(result, Matcher.quoteReplacement(replacement));
            } else {
                // Keep original if attachment not found
                matcher.appendReplacement(result, Matcher.quoteReplacement(matcher.group(0)));
            }
        }
        matcher.appendTail(result);

        return result.toString();
    }

    private static void appendEnhancedStyles(StringBuilder html) {
        int fontSize = StyleConstants.DEFAULT_FONT_SIZE;
        String textColor = StyleConstants.DEFAULT_TEXT_COLOR;
        String backgroundColor = StyleConstants.DEFAULT_BACKGROUND_COLOR;
        String borderColor = StyleConstants.DEFAULT_BORDER_COLOR;

        String css =
                String.format(
                        """
            body {
              font-family: %s;
              font-size: %dpx;
              line-height: %s;
              color: %s;
              margin: 0;
              padding: 16px;
              background-color: %s;
            }

            .email-container {
              width: 100%%;
              max-width: 100%%;
              margin: 0 auto;
            }

            .email-header {
              padding-bottom: 10px;
              border-bottom: 1px solid %s;
              margin-bottom: 10px;
            }

            .email-header h1 {
              margin: 0 0 10px 0;
              font-size: %dpx;
              font-weight: bold;
            }

            .email-meta div {
              margin-bottom: 2px;
              font-size: %dpx;
            }

            .email-body {
              word-wrap: break-word;
            }

            .attachment-section {
              margin-top: 15px;
              padding: 10px;
              background-color: %s;
              border: 1px solid %s;
              border-radius: 3px;
            }

            .attachment-section h3 {
              margin: 0 0 8px 0;
              font-size: %dpx;
            }

            .attachment-item {
              padding: 5px 0;
            }

            .attachment-icon {
              margin-right: 5px;
            }

            .attachment-details, .attachment-type {
              font-size: %dpx;
              color: #555555;
            }

            .attachment-inclusion-note, .attachment-info-note {
              margin-top: 8px;
              padding: 6px;
              font-size: %dpx;
              border-radius: 3px;
            }

            .attachment-inclusion-note {
              background-color: #e6ffed;
              border: 1px solid #d4f7dc;
              color: #006420;
            }

            .attachment-info-note {
              background-color: #fff9e6;
              border: 1px solid #fff0c2;
              color: #664d00;
            }

            .attachment-link-container {
              display: flex;
              align-items: center;
              padding: 8px;
              background-color: #f8f9fa;
              border: 1px solid #dee2e6;
              border-radius: 4px;
              margin: 4px 0;
            }

            .attachment-link-container:hover {
              background-color: #e9ecef;
            }

            .attachment-note {
              font-size: %dpx;
              color: #6c757d;
              font-style: italic;
              margin-left: 8px;
            }

            img {
              max-width: 100%%;
              height: auto;
              display: block;
            }
            """,
                        StyleConstants.DEFAULT_FONT_FAMILY,
                        fontSize,
                        StyleConstants.DEFAULT_LINE_HEIGHT,
                        textColor,
                        backgroundColor,
                        borderColor,
                        fontSize + 4,
                        fontSize - 1,
                        StyleConstants.ATTACHMENT_BACKGROUND_COLOR,
                        StyleConstants.ATTACHMENT_BORDER_COLOR,
                        fontSize + 1,
                        fontSize - 2,
                        fontSize - 2,
                        fontSize - 3);

        html.append(css);
    }

    private static String escapeHtml(String text) {
        if (text == null) return "";
        // Simple HTML escaping - for more complex sanitization, use CustomHtmlSanitizer
        return text.replace("&", "&amp;")
                .replace("<", "&lt;")
                .replace(">", "&gt;")
                .replace("\"", "&quot;")
                .replace("'", "&#39;");
    }

    private static String sanitizeText(String text, CustomHtmlSanitizer customHtmlSanitizer) {
        if (customHtmlSanitizer != null) {
            return customHtmlSanitizer.sanitize(text);
        } else {
            return escapeHtml(text);
        }
    }

    private static stirling.software.common.model.api.converters.HTMLToPdfRequest createHtmlRequest(
            EmlToPdfRequest request) {
        stirling.software.common.model.api.converters.HTMLToPdfRequest htmlRequest =
                new stirling.software.common.model.api.converters.HTMLToPdfRequest();

        if (request != null) {
            htmlRequest.setFileInput(request.getFileInput());
        }

        // Set default zoom level
        htmlRequest.setZoom(Float.parseFloat(StyleConstants.DEFAULT_ZOOM));

        return htmlRequest;
    }

    private static EmailContent extractEmailContentAdvanced(
            Object message, EmlToPdfRequest request, CustomHtmlSanitizer customHtmlSanitizer) {
        EmailContent content = new EmailContent();

        try {
            Class<?> messageClass = message.getClass();

            // Extract headers via reflection
            Method getSubject = messageClass.getMethod("getSubject");
            String subject = (String) getSubject.invoke(message);
            content.setSubject(subject != null ? safeMimeDecode(subject) : "No Subject");

            Method getFrom = messageClass.getMethod("getFrom");
            Object[] fromAddresses = (Object[]) getFrom.invoke(message);
            content.setFrom(
                    fromAddresses != null && fromAddresses.length > 0
                            ? safeMimeDecode(fromAddresses[0].toString())
                            : "");

            Method getAllRecipients = messageClass.getMethod("getAllRecipients");
            Object[] recipients = (Object[]) getAllRecipients.invoke(message);
            content.setTo(
                    recipients != null && recipients.length > 0
                            ? safeMimeDecode(recipients[0].toString())
                            : "");

            Method getSentDate = messageClass.getMethod("getSentDate");
            content.setDate((Date) getSentDate.invoke(message));

            // Extract content
            Method getContent = messageClass.getMethod("getContent");
            Object messageContent = getContent.invoke(message);

            if (messageContent instanceof String stringContent) {
                Method getContentType = messageClass.getMethod("getContentType");
                String contentType = (String) getContentType.invoke(message);
                if (contentType != null
                        && contentType.toLowerCase().contains(MimeConstants.TEXT_HTML)) {
                    content.setHtmlBody(stringContent);
                } else {
                    content.setTextBody(stringContent);
                }
            } else {
                // Handle multipart content
                try {
                    Class<?> multipartClass = Class.forName("jakarta.mail.Multipart");
                    if (multipartClass.isInstance(messageContent)) {
                        processMultipartAdvanced(
                                messageContent, content, request, customHtmlSanitizer);
                    }
                } catch (Exception e) {
                    log.warn("Error processing content: {}", e.getMessage());
                }
            }

        } catch (Exception e) {
            content.setSubject("Email Conversion");
            content.setFrom("Unknown");
            content.setTo("Unknown");
            content.setTextBody("Email content could not be parsed with advanced processing");
        }

        return content;
    }

    private static void processMultipartAdvanced(
            Object multipart,
            EmailContent content,
            EmlToPdfRequest request,
            CustomHtmlSanitizer customHtmlSanitizer) {
        try {
            // Enhanced multipart type checking
            if (!isValidJakartaMailMultipart(multipart)) {
                log.warn("Invalid Jakarta Mail multipart type: {}", multipart.getClass().getName());
                return;
            }

            Class<?> multipartClass = multipart.getClass();
            Method getCount = multipartClass.getMethod("getCount");
            int count = (Integer) getCount.invoke(multipart);

            Method getBodyPart = multipartClass.getMethod("getBodyPart", int.class);

            for (int i = 0; i < count; i++) {
                Object part = getBodyPart.invoke(multipart, i);
                processPartAdvanced(part, content, request, customHtmlSanitizer);
            }

        } catch (Exception e) {
            content.setTextBody("Email content could not be parsed with advanced processing");
        }
    }

    private static void processPartAdvanced(
            Object part,
            EmailContent content,
            EmlToPdfRequest request,
            CustomHtmlSanitizer customHtmlSanitizer) {
        try {
            if (!isValidJakartaMailPart(part)) {
                log.warn("Invalid Jakarta Mail part type: {}", part.getClass().getName());
                return;
            }

            Class<?> partClass = part.getClass();
            Method isMimeType = partClass.getMethod("isMimeType", String.class);
            Method getContent = partClass.getMethod("getContent");
            Method getDisposition = partClass.getMethod("getDisposition");
            Method getFileName = partClass.getMethod("getFileName");
            Method getContentType = partClass.getMethod("getContentType");
            Method getHeader = partClass.getMethod("getHeader", String.class);

            Object disposition = getDisposition.invoke(part);
            String filename = (String) getFileName.invoke(part);
            String contentType = (String) getContentType.invoke(part);

            if ((Boolean) isMimeType.invoke(part, MimeConstants.TEXT_PLAIN)
                    && disposition == null) {
                content.setTextBody((String) getContent.invoke(part));
            } else if ((Boolean) isMimeType.invoke(part, MimeConstants.TEXT_HTML)
                    && disposition == null) {
                String htmlBody = (String) getContent.invoke(part);
                // Apply sanitization to the HTML content if sanitizer is available
                if (customHtmlSanitizer != null) {
                    htmlBody = customHtmlSanitizer.sanitize(htmlBody);
                }
                content.setHtmlBody(htmlBody);
            } else if (MimeConstants.DISPOSITION_ATTACHMENT.equalsIgnoreCase((String) disposition)
                    || (filename != null && !filename.trim().isEmpty())) {

                content.setAttachmentCount(content.getAttachmentCount() + 1);

                // Always extract basic attachment metadata for display
                if (filename != null && !filename.trim().isEmpty()) {
                    // Create attachment with metadata only
                    EmailAttachment attachment = new EmailAttachment();
                    // Apply MIME decoding to filename to handle encoded attachment names
                    attachment.setFilename(safeMimeDecode(filename));
                    attachment.setContentType(contentType);

                    // Check if it's an embedded image
                    String[] contentIdHeaders =
                            (String[]) getHeader.invoke(part, MimeConstants.HEADER_CONTENT_ID);
                    if (contentIdHeaders != null && contentIdHeaders.length > 0) {
                        attachment.setEmbedded(true);
                        // Store the Content-ID, removing angle brackets if present
                        String contentId = contentIdHeaders[0];
                        if (contentId.startsWith("<") && contentId.endsWith(">")) {
                            contentId = contentId.substring(1, contentId.length() - 1);
                        }
                        attachment.setContentId(contentId);
                    }

                    // Extract attachment data if attachments should be included OR if it's an
                    // embedded image (needed for inline display)
                    if ((request != null && request.isIncludeAttachments())
                            || attachment.isEmbedded()) {
                        try {
                            Object attachmentContent = getContent.invoke(part);
                            byte[] attachmentData = null;

                            if (attachmentContent instanceof java.io.InputStream inputStream) {
                                try {
                                    attachmentData = inputStream.readAllBytes();
                                } catch (IOException e) {
                                    log.warn(
                                            "Failed to read InputStream attachment: {}",
                                            e.getMessage());
                                }
                            } else if (attachmentContent instanceof byte[] byteArray) {
                                attachmentData = byteArray;
                            } else if (attachmentContent instanceof String stringContent) {
                                attachmentData = stringContent.getBytes(StandardCharsets.UTF_8);
                            }

                            if (attachmentData != null) {
                                // Check size limit (use default 10MB if request is null)
                                long maxSizeMB =
                                        request != null ? request.getMaxAttachmentSizeMB() : 10L;
                                long maxSizeBytes = maxSizeMB * 1024 * 1024;

                                if (attachmentData.length <= maxSizeBytes) {
                                    attachment.setData(attachmentData);
                                    attachment.setSizeBytes(attachmentData.length);
                                } else {
                                    // For embedded images, always include data regardless of size
                                    // to ensure inline display works
                                    if (attachment.isEmbedded()) {
                                        attachment.setData(attachmentData);
                                        attachment.setSizeBytes(attachmentData.length);
                                    } else {
                                        // Still show attachment info even if too large
                                        attachment.setSizeBytes(attachmentData.length);
                                    }
                                }
                            }
                        } catch (Exception e) {
                            log.warn("Error extracting attachment data: {}", e.getMessage());
                        }
                    }

                    // Add attachment to the list for display (with or without data)
                    content.getAttachments().add(attachment);
                }
            } else if ((Boolean) isMimeType.invoke(part, MimeConstants.MULTIPART_PREFIX + "*")) {
                // Handle nested multipart content
                try {
                    Object multipartContent = getContent.invoke(part);
                    Class<?> multipartClass = Class.forName("jakarta.mail.Multipart");
                    if (multipartClass.isInstance(multipartContent)) {
                        processMultipartAdvanced(
                                multipartContent, content, request, customHtmlSanitizer);
                    }
                } catch (Exception e) {
                    log.warn("Error processing multipart content: {}", e.getMessage());
                }
            }

        } catch (Exception e) {
            log.warn("Error processing multipart part: {}", e.getMessage());
        }
    }

    private static String generateEnhancedEmailHtml(
            EmailContent content,
            EmlToPdfRequest request,
            CustomHtmlSanitizer customHtmlSanitizer) {
        StringBuilder html = new StringBuilder();

        html.append(
                String.format(
                        """
            <!DOCTYPE html>
            <html><head><meta charset="UTF-8">
            <title>%s</title>
            <style>
            """,
                        sanitizeText(content.getSubject(), customHtmlSanitizer)));
        appendEnhancedStyles(html);
        html.append(
                """
            </style>
            </head><body>
            """);

        html.append(
                String.format(
                        """
            <div class="email-container">
            <div class="email-header">
            <h1>%s</h1>
            <div class="email-meta">
            <div><strong>From:</strong> %s</div>
            <div><strong>To:</strong> %s</div>
            """,
                        sanitizeText(content.getSubject(), customHtmlSanitizer),
                        sanitizeText(content.getFrom(), customHtmlSanitizer),
                        sanitizeText(content.getTo(), customHtmlSanitizer)));

        if (content.getDate() != null) {
            html.append(
                    String.format(
                            "<div><strong>Date:</strong> %s</div>\n",
                            formatEmailDate(content.getDate())));
        }
        html.append(
                """
            </div></div>
            """);

        html.append("<div class=\"email-body\">\n");
        if (content.getHtmlBody() != null && !content.getHtmlBody().trim().isEmpty()) {
            html.append(processEmailHtmlBody(content.getHtmlBody(), content, customHtmlSanitizer));
        } else if (content.getTextBody() != null && !content.getTextBody().trim().isEmpty()) {
            html.append(
                    String.format(
                            """
                <div class="text-body">%s</div>""",
                            convertTextToHtml(content.getTextBody(), customHtmlSanitizer)));
        } else {
            html.append(
                    """
                <div class="no-content">
                <p><em>No content available</em></p>
                </div>""");
        }
        html.append("</div>\n");

        if (content.getAttachmentCount() > 0 || !content.getAttachments().isEmpty()) {
            html.append("<div class=\"attachment-section\">\n");
            int displayedAttachmentCount =
                    content.getAttachmentCount() > 0
                            ? content.getAttachmentCount()
                            : content.getAttachments().size();
            html.append("<h3>Attachments (").append(displayedAttachmentCount).append(")</h3>\n");

            if (!content.getAttachments().isEmpty()) {
                for (EmailAttachment attachment : content.getAttachments()) {
                    // Create attachment info with paperclip emoji before filename
                    String uniqueId = generateUniqueAttachmentId(attachment.getFilename());
                    attachment.setEmbeddedFilename(
                            attachment.getEmbeddedFilename() != null
                                    ? attachment.getEmbeddedFilename()
                                    : attachment.getFilename());

                    String sizeStr = formatFileSize(attachment.getSizeBytes());
                    String contentType =
                            attachment.getContentType() != null
                                            && !attachment.getContentType().isEmpty()
                                    ? ", " + escapeHtml(attachment.getContentType())
                                    : "";

                    html.append(
                            String.format(
                                    """
                        <div class="attachment-item" id="%s">
                        <span class="attachment-icon">%s</span>
                        <span class="attachment-name">%s</span>
                        <span class="attachment-details">(%s%s)</span>
                        </div>
                        """,
                                    uniqueId,
                                    MimeConstants.ATTACHMENT_MARKER,
                                    escapeHtml(safeMimeDecode(attachment.getFilename())),
                                    sizeStr,
                                    contentType));
                }
            }

            if (request.isIncludeAttachments()) {
                html.append(
                        """
                    <div class="attachment-info-note">
                    <p><em>Attachments are embedded in the file.</em></p>
                    </div>
                    """);
            } else {
                html.append(
                        """
                    <div class="attachment-info-note">
                    <p><em>Attachment information displayed - files not included in PDF.</em></p>
                    </div>
                    """);
            }

            html.append("</div>\n");
        }

        html.append(
                """
            </div>
            </body></html>""");

        return html.toString();
    }

    private static byte[] attachFilesToPdf(
            byte[] pdfBytes,
            List<EmailAttachment> attachments,
            CustomPDFDocumentFactory pdfDocumentFactory)
            throws IOException {

        if (attachments == null || attachments.isEmpty()) {
            return pdfBytes;
        }

        try (PDDocument document = pdfDocumentFactory.load(pdfBytes);
                ByteArrayOutputStream outputStream = new ByteArrayOutputStream()) {

            // Convert EmailAttachment objects to MultipartFile objects
            List<MultipartFile> multipartAttachments = new ArrayList<>();
            for (EmailAttachment attachment : attachments) {
                if (attachment.getData() != null && attachment.getData().length > 0) {
                    multipartAttachments.add(createMultipartFile(attachment));
                }
            }

            if (!multipartAttachments.isEmpty()) {
                // Use existing AttachmentService-style logic for embedding files
                addAttachmentsToDocument(document, multipartAttachments);

                // Set catalog viewer preferences to show attachments pane
                setCatalogViewerPreferences(document, PageMode.USE_ATTACHMENTS);

                // Add attachment annotations for EML-specific functionality
                addAttachmentAnnotationsToDocument(document, attachments);
            }

            document.save(outputStream);
            return outputStream.toByteArray();
        }
    }

    private static MultipartFile createMultipartFile(EmailAttachment attachment) {
        return new MultipartFile() {
            @Override
            public @NotNull String getName() {
                return "attachment";
            }

            @Override
            public String getOriginalFilename() {
                return attachment.getFilename() != null
                        ? attachment.getFilename()
                        : "attachment_" + System.currentTimeMillis();
            }

            @Override
            public String getContentType() {
                return attachment.getContentType() != null
                        ? attachment.getContentType()
                        : "application/octet-stream";
            }

            @Override
            public boolean isEmpty() {
                return attachment.getData() == null || attachment.getData().length == 0;
            }

            @Override
            public long getSize() {
                return attachment.getData() != null ? attachment.getData().length : 0;
            }

            @Override
            public byte @NotNull [] getBytes() {
                return attachment.getData();
            }

            @Override
            public @NotNull InputStream getInputStream() {
                return new ByteArrayInputStream(attachment.getData());
            }

            @Override
            public void transferTo(@NotNull File dest) throws IOException, IllegalStateException {
                try (FileOutputStream fos = new java.io.FileOutputStream(dest)) {
                    fos.write(attachment.getData());
                }
            }
        };
    }

    private static void addAttachmentsToDocument(
            PDDocument document, List<MultipartFile> attachments) throws IOException {
        // Get or create embedded files tree
        PDDocumentCatalog catalog = document.getDocumentCatalog();
        PDDocumentNameDictionary documentNames = catalog.getNames();
        if (documentNames == null) {
            documentNames = new PDDocumentNameDictionary(catalog);
            catalog.setNames(documentNames);
        }

        PDEmbeddedFilesNameTreeNode embeddedFilesTree = documentNames.getEmbeddedFiles();
        if (embeddedFilesTree == null) {
            embeddedFilesTree = new PDEmbeddedFilesNameTreeNode();
            documentNames.setEmbeddedFiles(embeddedFilesTree);
        }

        Map<String, PDComplexFileSpecification> existingNames;
        try {
            Map<String, PDComplexFileSpecification> names = embeddedFilesTree.getNames();
            existingNames = names != null ? new HashMap<>(names) : new HashMap<>();
        } catch (IOException e) {
            log.error("Could not retrieve existing embedded files", e);
            throw e;
        }

        // Add each attachment using simplified logic from AttachmentService
        for (MultipartFile attachment : attachments) {
            String filename = attachment.getOriginalFilename();
            if (filename == null || filename.trim().isEmpty()) {
                filename = "attachment_" + System.currentTimeMillis();
            }

            try {
                PDEmbeddedFile embeddedFile =
                        new PDEmbeddedFile(document, attachment.getInputStream());
                embeddedFile.setSize((int) attachment.getSize());
                embeddedFile.setCreationDate(new GregorianCalendar());
                embeddedFile.setModDate(new GregorianCalendar());

                String contentType = attachment.getContentType();
                if (contentType != null && !contentType.trim().isEmpty()) {
                    embeddedFile.setSubtype(contentType);
                }

                PDComplexFileSpecification fileSpecification = new PDComplexFileSpecification();
                fileSpecification.setFile(filename);
                fileSpecification.setFileUnicode(filename);
                fileSpecification.setFileDescription("Email attachment: " + filename);
                fileSpecification.setEmbeddedFile(embeddedFile);
                fileSpecification.setEmbeddedFileUnicode(embeddedFile);

                existingNames.put(filename, fileSpecification);
                log.info("Added email attachment: {} ({} bytes)", filename, attachment.getSize());
            } catch (IOException e) {
                log.warn("Failed to create embedded file for attachment: {}", filename, e);
            }
        }

        embeddedFilesTree.setNames(existingNames);
    }

    private static void addAttachmentAnnotationsToDocument(
            PDDocument document, List<EmailAttachment> attachments) throws IOException {
        if (document.getNumberOfPages() == 0 || attachments == null || attachments.isEmpty()) {
            return;
        }

        // 1. Find the screen position of all attachment markers
        AttachmentMarkerPositionFinder finder = new AttachmentMarkerPositionFinder();
        finder.setSortByPosition(true); // Process pages in order
        finder.getText(document);
        List<MarkerPosition> markerPositions = finder.getPositions();

        // 2. Warn if the number of markers and attachments don't match
        if (markerPositions.size() != attachments.size()) {
            log.warn(
                    "Found {} attachment markers, but there are {} attachments. Annotation count may be incorrect.",
                    markerPositions.size(),
                    attachments.size());
        }

        // 3. Create an invisible annotation over each found marker
        int annotationsToAdd = Math.min(markerPositions.size(), attachments.size());
        for (int i = 0; i < annotationsToAdd; i++) {
            MarkerPosition position = markerPositions.get(i);
            EmailAttachment attachment = attachments.get(i);

            if (attachment.getEmbeddedFilename() != null) {
                PDPage page = document.getPage(position.getPageIndex());
                addAttachmentAnnotationToPage(
                        document, page, attachment, position.getX(), position.getY());
            }
        }
    }

    private static void addAttachmentAnnotationToPage(
            PDDocument document, PDPage page, EmailAttachment attachment, float x, float y)
            throws IOException {

        PDAnnotationFileAttachment fileAnnotation = new PDAnnotationFileAttachment();

        PDRectangle rect = getPdRectangle(page, x, y);
        fileAnnotation.setRectangle(rect);

        // Remove visual appearance while keeping clickable functionality
        try {
            PDAppearanceDictionary appearance = new PDAppearanceDictionary();
            PDAppearanceStream normalAppearance = new PDAppearanceStream(document);
            normalAppearance.setBBox(new PDRectangle(0, 0, 0, 0)); // Zero-size bounding box

            appearance.setNormalAppearance(normalAppearance);
            fileAnnotation.setAppearance(appearance);
        } catch (Exception e) {
            // If appearance manipulation fails, just set it to null
            fileAnnotation.setAppearance(null);
        }

        // Set invisibility flags but keep it functional
        fileAnnotation.setInvisible(true);
        fileAnnotation.setHidden(false); // Must be false to remain clickable
        fileAnnotation.setNoView(false); // Must be false to remain clickable
        fileAnnotation.setPrinted(false);

        PDEmbeddedFilesNameTreeNode efTree =
                document.getDocumentCatalog().getNames().getEmbeddedFiles();
        if (efTree != null) {
            Map<String, PDComplexFileSpecification> efMap = efTree.getNames();
            if (efMap != null) {
                PDComplexFileSpecification fileSpec = efMap.get(attachment.getEmbeddedFilename());
                if (fileSpec != null) {
                    fileAnnotation.setFile(fileSpec);
                }
            }
        }

        fileAnnotation.setContents("Click to open: " + attachment.getFilename());
        fileAnnotation.setAnnotationName("EmbeddedFile_" + attachment.getEmbeddedFilename());

        page.getAnnotations().add(fileAnnotation);

        log.info(
                "Added attachment annotation for '{}' on page {}",
                attachment.getFilename(),
                document.getPages().indexOf(page) + 1);
    }

    private static @NotNull PDRectangle getPdRectangle(PDPage page, float x, float y) {
        PDRectangle mediaBox = page.getMediaBox();
        float pdfY = mediaBox.getHeight() - y;

        float iconWidth =
                StyleConstants.ATTACHMENT_ICON_WIDTH; // Keep original size for clickability
        float iconHeight =
                StyleConstants.ATTACHMENT_ICON_HEIGHT; // Keep original size for clickability

        // Keep the full-size rectangle so it remains clickable
        return new PDRectangle(
                x + StyleConstants.ANNOTATION_X_OFFSET,
                pdfY - iconHeight + StyleConstants.ANNOTATION_Y_OFFSET,
                iconWidth,
                iconHeight);
    }

    private static String formatEmailDate(Date date) {
        if (date == null) return "";
        java.text.SimpleDateFormat formatter =
                new java.text.SimpleDateFormat("EEE, MMM d, yyyy 'at' h:mm a", Locale.ENGLISH);
        return formatter.format(date);
    }

    private static String formatFileSize(long bytes) {
        return GeneralUtils.formatBytes(bytes);
    }

    /**
     * Safely decode MIME headers using Jakarta Mail if available, fallback to custom implementation
     */
    private static String safeMimeDecode(String headerValue) {
        if (headerValue == null || headerValue.trim().isEmpty()) {
            return "";
        }

        // Initialize MimeUtility decoding if not already checked
        if (!mimeUtilityChecked) {
            initializeMimeUtilityDecoding();
        }

        // Try Jakarta Mail MimeUtility.decodeText first if available
        if (mimeUtilityDecodeTextMethod != null) {
            try {
                return (String) mimeUtilityDecodeTextMethod.invoke(null, headerValue.trim());
            } catch (Exception e) {
                log.debug(
                        "Failed to decode MIME header via MimeUtility, using fallback: {}",
                        headerValue);
            }
        }

        // Fallback to custom decoding
        return decodeMimeHeader(headerValue.trim());
    }

    private static void initializeMimeUtilityDecoding() {
        try {
            Class<?> mimeUtilityClass = Class.forName("jakarta.mail.internet.MimeUtility");
            mimeUtilityDecodeTextMethod = mimeUtilityClass.getMethod("decodeText", String.class);
            log.debug("Jakarta Mail MimeUtility.decodeText method available");
        } catch (Exception e) {
            log.debug("Jakarta Mail MimeUtility.decodeText method not available, using fallback");
            mimeUtilityDecodeTextMethod = null;
        }
        mimeUtilityChecked = true;
    }

    private static String decodeMimeHeader(String encodedText) {
        if (encodedText == null || encodedText.trim().isEmpty()) {
            return encodedText;
        }

        try {
            StringBuilder result = new StringBuilder();
            Matcher matcher = MimeConstants.MIME_ENCODED_PATTERN.matcher(encodedText);
            int lastEnd = 0;

            while (matcher.find()) {
                // Add any text before the encoded part
                result.append(encodedText, lastEnd, matcher.start());

                String charset = matcher.group(1);
                String encoding = matcher.group(2).toUpperCase();
                String encodedValue = matcher.group(3);

                try {
                    String decodedValue;
                    if ("B".equals(encoding)) {
                        // Base64 decoding
                        byte[] decodedBytes = Base64.getDecoder().decode(encodedValue);
                        decodedValue = new String(decodedBytes, Charset.forName(charset));
                    } else if ("Q".equals(encoding)) {
                        // Quoted-printable decoding
                        decodedValue = decodeQuotedPrintable(encodedValue, charset);
                    } else {
                        // Unknown encoding, keep original
                        decodedValue = matcher.group(0);
                    }
                    result.append(decodedValue);
                } catch (Exception e) {
                    log.warn("Failed to decode MIME header part: {}", matcher.group(0), e);
                    // If decoding fails, keep the original encoded text
                    result.append(matcher.group(0));
                }

                lastEnd = matcher.end();
            }

            // Add any remaining text after the last encoded part
            result.append(encodedText.substring(lastEnd));

            return result.toString();
        } catch (Exception e) {
            log.warn("Error decoding MIME header: {}", encodedText, e);
            return encodedText; // Return original if decoding fails
        }
    }

    private static String decodeQuotedPrintable(String encodedText, String charset) {
        StringBuilder result = new StringBuilder();
        for (int i = 0; i < encodedText.length(); i++) {
            char c = encodedText.charAt(i);
            switch (c) {
                case '=' -> {
                    if (i + 2 < encodedText.length()) {
                        String hex = encodedText.substring(i + 1, i + 3);
                        try {
                            int value = Integer.parseInt(hex, 16);
                            result.append((char) value);
                            i += 2; // Skip the hex digits
                        } catch (NumberFormatException e) {
                            // If hex parsing fails, keep the original character
                            result.append(c);
                        }
                    } else {
                        result.append(c);
                    }
                }
                case '_' -> // In RFC 2047, underscore represents space
                        result.append(' ');
                default -> result.append(c);
            }
        }

        // Convert bytes to proper charset
        byte[] bytes = result.toString().getBytes(StandardCharsets.ISO_8859_1);
        return new String(bytes, Charset.forName(charset));
    }

    private static boolean isValidJakartaMailPart(Object part) {
        if (part == null) return false;

        try {
            // Check if the object implements jakarta.mail.Part interface
            Class<?> partInterface = Class.forName("jakarta.mail.Part");
            if (!partInterface.isInstance(part)) {
                return false;
            }

            // Additional check for MimePart
            try {
                Class<?> mimePartInterface = Class.forName("jakarta.mail.internet.MimePart");
                return mimePartInterface.isInstance(part);
            } catch (ClassNotFoundException e) {
                // MimePart not available, but Part is enough
                return true;
            }
        } catch (ClassNotFoundException e) {
            log.debug("Jakarta Mail Part interface not available for validation");
            return false;
        }
    }

    private static boolean isValidJakartaMailMultipart(Object multipart) {
        if (multipart == null) return false;

        try {
            // Check if the object implements jakarta.mail.Multipart interface
            Class<?> multipartInterface = Class.forName("jakarta.mail.Multipart");
            if (!multipartInterface.isInstance(multipart)) {
                return false;
            }

            // Additional check for MimeMultipart
            try {
                Class<?> mimeMultipartClass = Class.forName("jakarta.mail.internet.MimeMultipart");
                if (mimeMultipartClass.isInstance(multipart)) {
                    log.debug("Found MimeMultipart instance for enhanced processing");
                    return true;
                }
            } catch (ClassNotFoundException e) {
                log.debug("MimeMultipart not available, using base Multipart interface");
            }

            return true;
        } catch (ClassNotFoundException e) {
            log.debug("Jakarta Mail Multipart interface not available for validation");
            return false;
        }
    }

    @Data
    public static class EmailContent {
        private String subject;
        private String from;
        private String to;
        private Date date;
        private String htmlBody;
        private String textBody;
        private int attachmentCount;
        private List<EmailAttachment> attachments = new ArrayList<>();

        public void setHtmlBody(String htmlBody) {
            this.htmlBody = htmlBody != null ? htmlBody.replaceAll("\r", "") : null;
        }

        public void setTextBody(String textBody) {
            this.textBody = textBody != null ? textBody.replaceAll("\r", "") : null;
        }
    }

    @Data
    public static class EmailAttachment {
        private String filename;
        private String contentType;
        private byte[] data;
        private boolean embedded;
        private String embeddedFilename;
        private long sizeBytes;

        // New fields for advanced processing
        private String contentId;
        private String disposition;
        private String transferEncoding;

        // Custom setter to maintain size calculation logic
        public void setData(byte[] data) {
            this.data = data;
            if (data != null) {
                this.sizeBytes = data.length;
            }
        }
    }

    @Data
    public static class MarkerPosition {
        private int pageIndex;
        private float x;
        private float y;
        private String character;

        public MarkerPosition(int pageIndex, float x, float y, String character) {
            this.pageIndex = pageIndex;
            this.x = x;
            this.y = y;
            this.character = character;
        }
    }

    public static class AttachmentMarkerPositionFinder
            extends org.apache.pdfbox.text.PDFTextStripper {
        @Getter private final List<MarkerPosition> positions = new ArrayList<>();
        private int currentPageIndex;
        protected boolean sortByPosition;
        private boolean isInAttachmentSection;
        private boolean attachmentSectionFound;

        public AttachmentMarkerPositionFinder() {
            super();
            this.currentPageIndex = 0;
            this.sortByPosition = false;
            this.isInAttachmentSection = false;
            this.attachmentSectionFound = false;
        }

        @Override
        protected void startPage(org.apache.pdfbox.pdmodel.PDPage page) throws IOException {
            super.startPage(page);
        }

        @Override
        protected void endPage(org.apache.pdfbox.pdmodel.PDPage page) throws IOException {
            currentPageIndex++;
            super.endPage(page);
        }

        @Override
        protected void writeString(
                String string, List<org.apache.pdfbox.text.TextPosition> textPositions)
                throws IOException {
            String lowerString = string.toLowerCase();

            // Look for attachment section start marker
            if (lowerString.contains("attachments (")) {
                isInAttachmentSection = true;
                attachmentSectionFound = true;
            }

            // Look for attachment section end markers (common patterns that indicate end of
            // attachments)
            if (isInAttachmentSection
                    && (lowerString.contains("</body>")
                            || lowerString.contains("</html>")
                            || (attachmentSectionFound
                                    && lowerString.trim().isEmpty()
                                    && string.length() > 50))) {
                isInAttachmentSection = false;
            }

            // Only look for markers if we are in the attachment section
            if (isInAttachmentSection) {
                String attachmentMarker = MimeConstants.ATTACHMENT_MARKER;
                for (int i = 0; (i = string.indexOf(attachmentMarker, i)) != -1; i++) {
                    if (i < textPositions.size()) {
                        org.apache.pdfbox.text.TextPosition textPosition = textPositions.get(i);
                        MarkerPosition position =
                                new MarkerPosition(
                                        currentPageIndex,
                                        textPosition.getXDirAdj(),
                                        textPosition.getYDirAdj(),
                                        attachmentMarker);
                        positions.add(position);
                    }
                }
            }
            super.writeString(string, textPositions);
        }

        @Override
        public void setSortByPosition(boolean sortByPosition) {
            this.sortByPosition = sortByPosition;
        }
    }
}
