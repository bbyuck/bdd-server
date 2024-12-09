package com.bb.bdd.domain.excel.util;

import com.bb.bdd.domain.excel.DownloadCode;
import jakarta.servlet.http.HttpServletResponse;
import lombok.Getter;
import lombok.RequiredArgsConstructor;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.Workbook;
import org.springframework.beans.factory.annotation.Value;
import org.springframework.http.HttpHeaders;
import org.springframework.stereotype.Component;
import org.springframework.web.multipart.MultipartFile;

import java.io.*;
import java.net.URLEncoder;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.time.LocalDate;
import java.util.zip.ZipEntry;
import java.util.zip.ZipOutputStream;

@Slf4j
@Component
@RequiredArgsConstructor
public class FileManager {

    @Value("${file.temp}")
    private String tempFilePath;

    private final HttpServletResponse response;

    private void setDownloadResponseHeader(DownloadCode downloadCode, boolean isMulti) {
        if (isMulti) {
            response.setContentType("application/zip");
            response.setHeader(HttpHeaders.CONTENT_DISPOSITION, String.format(
                    "attachment; filename*=UTF-8'''%s_%s.%s'",
                    LocalDate.now(),
                    URLEncoder.encode(downloadCode.getFilename(), StandardCharsets.UTF_8), // 한글 파일명 꺠짐을 방지하기 위해 파일명을 UTF-8로 인코딩 해 헤더 set
                    "zip"
            ));
        } else {
            response.setContentType("application/octet-stream");
            response.setHeader(HttpHeaders.CONTENT_DISPOSITION, String.format(
                    "attachment; filename*=UTF-8'''%s_%s.%s'",
                    LocalDate.now(),
                    URLEncoder.encode(downloadCode.getFilename(), StandardCharsets.UTF_8), // 한글 파일명 꺠짐을 방지하기 위해 파일명을 UTF-8로 인코딩 해 헤더 set
                    "xlsx"
            ));
        }
        response.setHeader("Access-Control-Expose-Headers", "Content-Disposition");
    }

    public File createFileWithMultipartFile(MultipartFile multipartFile) {
        String tempFilePath = this.tempFilePath + File.separator + multipartFile.getOriginalFilename();
        File tempFile = new File(tempFilePath);

        try (InputStream is = multipartFile.getInputStream();
             OutputStream fos = new FileOutputStream(tempFile)) {

            byte[] buffer = new byte[1024];
            int readBytes;

            while ((readBytes = is.read(buffer)) != -1) {
                fos.write(buffer, 0, readBytes);
            }
            fos.flush();
        } catch (IOException e) {
            log.error(e.getMessage(), e);
            throw new RuntimeException("임시 파일을 생성하는 중 에러가 발생했습니다.");
        }

        return tempFile;
    }

    public File createTempFile(Workbook wb, String filePath) {
        File xlsFile = new File(tempFilePath + File.separator + filePath);

        try (FileOutputStream fileOut = new FileOutputStream(xlsFile)) {
            wb.write(fileOut);
        } catch (IOException e) {
            log.error(e.getMessage(), e);
            throw new RuntimeException("xls 임시 파일을 생성하는 과저엥서 문제가 발생했습니다.");
        }

        return xlsFile;
    }

    public File createTempFile(String extension) {
        return new File(tempFilePath + File.separator + System.currentTimeMillis() + "." + extension);
    }

    public void download(DownloadCode downloadCode, File... files) {
        boolean isMulti = files.length > 1;

        setDownloadResponseHeader(downloadCode, isMulti);

        if (isMulti) {
            compressAndDownloadStreaming(files);
        } else if (files.length == 1) {
            singleDownloadStreaming(files[0]);
        } else {
            log.error("잘못된 파일 다운로드 요청입니다.");
            throw new RuntimeException("잘못된 파일 다운로드 요청입니다.");
        }

    }

    private void singleDownloadStreaming(File file) {
        try (FileInputStream fis = new FileInputStream(file);
             OutputStream os = response.getOutputStream()) {
            byte[] buffer = new byte[1024];
            int readBytes;

            while((readBytes = fis.read(buffer)) != -1) {
                os.write(buffer, 0, readBytes);
            }

            os.flush();
        } catch (IOException e) {
            log.error(e.getMessage(), e);
            throw new RuntimeException("다운로드 받기 위한 파일을 읽어오는 도중 에러가 발생했습니다.");
        }
    }

    private void compressAndDownloadStreaming(File[] files) {
        try (ZipOutputStream zos = new ZipOutputStream(response.getOutputStream())) {
            for (File file : files) {
                ZipEntry entry = new ZipEntry(file.getName());
                zos.putNextEntry(entry);

                Files.copy(file.toPath(), zos);

                zos.closeEntry();
            }

            zos.finish();
        } catch (IOException e) {
            log.error(e.getMessage(), e);
            throw new RuntimeException("파일을 압축해서 다운로드하는 도중 문제가 발생했습니다.");
        }
    }

}
