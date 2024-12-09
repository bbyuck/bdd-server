package com.bb.bdd.domain.excel.service;

import com.bb.bdd.domain.excel.DownloadCode;
import com.bb.bdd.domain.excel.ShopCode;
import com.bb.bdd.domain.excel.dto.CnpInputDto;
import com.bb.bdd.domain.excel.dto.CoupangColumnDto;
import com.bb.bdd.domain.excel.dto.NaverColumnDto;
import com.bb.bdd.domain.excel.dto.Pair;
import com.bb.bdd.domain.excel.util.ExcelReader;
import com.bb.bdd.domain.excel.util.FileManager;
import jakarta.annotation.PostConstruct;
import lombok.RequiredArgsConstructor;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Service;
import org.springframework.util.StringUtils;
import org.springframework.web.multipart.MultipartFile;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.util.*;
import java.util.Map.Entry;

@Slf4j
@Service
@RequiredArgsConstructor
public class ExcelTransformService {

    @PostConstruct
    public void init() throws IOException {
        readCoupangProducts();
        readNaverProducts();
    }

    private final HashMap<String, String> coupangDict = new HashMap<>();
    private final HashMap<String, String> naverProductDict = new HashMap<>();
    private final HashMap<String, String> naverOptionDict = new HashMap<>();

    // 쿠팡 제목 바꾸는 엑셀파일
    private final String COUPANG_DICT_URI = "dict/coupang_dict.xlsx";
    private final String NAVER_DICT_URI = "dict/naver_dict.xlsx";
    private final Integer COUPANG_OPTION_START_IDX = 10;

    private final String COURIER = "대한통운";

    private final ExcelReader excelReader;
    private final FileManager fileManager;

    private static final ThreadLocal<Map<String, Integer>> threadLocalCoupangCountMap = ThreadLocal.withInitial(() -> null);
    private static final ThreadLocal<Map<String, Integer>> threadLocalNaverCountMap = ThreadLocal.withInitial(() -> null);

    private DateTimeFormatter formatter = DateTimeFormatter.ofPattern("yyyy-MM-dd-HH-mm-ss");

    /**
     * ====================================  공통  ========================================
     */


    /**
     * 주문 목록 엑셀 파일을 읽어 CNP에 업로드할 수 있는 양식의 엑셀 파일을 생성한다.
     *
     * @param excelFile
     * @return
     */
    public File createCnpXls(File excelFile, ShopCode shopCode) {
        String tempFileName = LocalDateTime.now().format(formatter) + "_" + shopCode.getValue() + "- Cnp.xls";

        // cnp input list
        List<CnpInputDto> cnpInputLs =
                shopCode == ShopCode.COUPANG
                        ? parseCoupang(excelFile)
                        : shopCode == ShopCode.NAVER
                        ? parseNaver(excelFile) : null;

        try (HSSFWorkbook xlsWb = new HSSFWorkbook()) {
            // sheet 생성
            HSSFSheet sheet = xlsWb.createSheet("배송관리");

            // 스타일
            CellStyle menu = xlsWb.createCellStyle();
            CellStyle defaultStyle = xlsWb.createCellStyle();


            // 줄바꿈
            menu.setWrapText(true);
            defaultStyle.setWrapText(true);

            // 메뉴
            HSSFRow row = sheet.createRow(0);
            HSSFCell cell = row.createCell(0);

            cell.setCellValue("수취인이름");

            cell = row.createCell(1);
            cell.setCellValue("수취인전화번호");
            cell = row.createCell(2);
            cell.setCellValue("수취인 주소");
            cell = row.createCell(3);
            cell.setCellValue("노출상품명(옵션명)");
            cell = row.createCell(4);
            cell.setCellValue("비고");

            for (int i = 1; i <= cnpInputLs.size(); i++) {
                CnpInputDto cnpInput = cnpInputLs.get(i - 1);
                row = sheet.createRow(i);

                cell = row.createCell(0);
                cell.setCellValue(cnpInput.getReceiverName());

                cell = row.createCell(1);
                cell.setCellValue(cnpInput.getReceiverPhone());

                cell = row.createCell(2);
                cell.setCellValue(cnpInput.getReceiverAddress());

                cell = row.createCell(3);
                cell.setCellValue(cnpInput.getOrderContents());

                cell = row.createCell(4);
                cell.setCellValue(cnpInput.getRemark());
            }

            return fileManager.createTempFile(xlsWb, tempFileName);
        } catch (Exception e) {
            log.error(e.getMessage(), e);
            throw new RuntimeException("쿠팡 운송장 출력 엑셀 파일을 생성하는 중 에러가 발생했습니다.");
        }
    }

    /**
     * 주문 목록 엑셀 파일을 읽어 판매량 집계 엑셀 파일을 생성한다.
     *
     * @param excelFile, shopCode
     * @return
     */
    public File createCountXlsx(File excelFile, ShopCode shopCode) {
        try (XSSFWorkbook xlsxWb = new XSSFWorkbook()) {
            List<Pair> countLs = shopCode == ShopCode.COUPANG ? coupangCountList() : shopCode == ShopCode.NAVER ? naverCountList() : null;
            String tempFileName = LocalDateTime.now().format(formatter) + "_" + shopCode.getValue() + " 판매량.xlsx";

            // sheet 생성
            XSSFSheet sheet = xlsxWb.createSheet(LocalDateTime.now().format(formatter) + " 판매량");

            // 스타일
            CellStyle menu = xlsxWb.createCellStyle();
            CellStyle defaultStyle = xlsxWb.createCellStyle();

            // 줄바꿈
            menu.setWrapText(true);
            defaultStyle.setWrapText(true);

            // 메뉴
            XSSFRow row = sheet.createRow(0);
            XSSFCell cell = row.createCell(0);
            cell.setCellValue("판매 물품 이름");
            row.createCell(1);
            cell.setCellValue("판매량");

            for (int i = 1; i <= countLs.size(); i++) {
                Pair pair = countLs.get(i - 1);
                row = sheet.createRow(i);

                cell = row.createCell(0);
                cell.setCellValue(pair.getProductName());
                cell = row.createCell(1);
                cell.setCellValue(pair.getCount());
            }

            return fileManager.createTempFile(xlsxWb, tempFileName);
        } catch (Exception e) {
            log.error(e.getMessage(), e);
            throw new RuntimeException(shopCode.getValue() + " 판매량 집계 엑셆 파일을 생성하는 중 에러가 발생했습니다.");
        } finally {
            threadLocalNaverCountMap.remove();
            threadLocalCoupangCountMap.remove();
        }
    }

    /**
     * 주문 목록 파일을 MultiparFile로 받아 임시파일을 생성하고
     * ShopCode에 따라 아래 목록에 대한 엑셀 파일을 생성해 다운로드 처리한다.
     *
     * <p>
     * - 주문 목록에 대한 운송장 번호 출력을 위한 업로드 엑셀
     * - 주문량 집계
     *
     * @param multipartFile
     */
    public void transformToCnp(MultipartFile multipartFile, ShopCode shopCode) {
        File tempExcelFile = null;
        File cnpXlsTempFile = null;
        File countXlsxTempFile = null;

        try {
            tempExcelFile = fileManager.createFileWithMultipartFile(multipartFile);
            cnpXlsTempFile = createCnpXls(tempExcelFile, shopCode);
            countXlsxTempFile = createCountXlsx(tempExcelFile, shopCode);

            fileManager.download(
                    shopCode == ShopCode.COUPANG ? DownloadCode.COUPANG_CNP
                            : shopCode == ShopCode.NAVER ? DownloadCode.NAVER_CNP
                            : null
                    , cnpXlsTempFile, countXlsxTempFile);
        } finally {
            tempExcelFile.delete();
            cnpXlsTempFile.delete();
            countXlsxTempFile.delete();
        }
    }

    /**
     * ====================================  공통  ========================================
     */


    /**
     * ====================================  쿠팡  ========================================
     */

    /**
     * 빈 생성 시점에 쿠팡 판매목록 엑셀을 읽어 메모리에 로드한다.
     */
    private void readCoupangProducts() {
        XSSFWorkbook workbook = excelReader.readXlsxOnClassPath(COUPANG_DICT_URI);

        XSSFSheet sheet = workbook.getSheetAt(0); // 해당 엑셀파일의 시트수
        int rows = sheet.getPhysicalNumberOfRows(); // 해당 시트의 행의 개수

        // 파싱
        for (int rowIdx = 1; rowIdx < rows; rowIdx++) {
            XSSFRow row = sheet.getRow(rowIdx); // 각 행을 읽어온다

            if (row != null) {
                XSSFCell cell = row.getCell(6);
                String key = cell.getStringCellValue();
                cell = row.getCell(8);
                String value = cell.getStringCellValue();
                coupangDict.put(key, value);
            }
        }
    }

    /**
     * 현재 요청에서 집계된 쿠팡 판매량 집계 Map을 판매목록 엑셀을 읽어 생성하거나
     * Thread safe한 HttpServletRequest의 attribute에서 가져온다.
     */
    private Map<String, Integer> getCoupangCountMap() {
        if (threadLocalCoupangCountMap.get() != null) {
            return threadLocalCoupangCountMap.get();
        }

        Map<String, Integer> coupangCountMap = new HashMap<>();

        try (XSSFWorkbook wb = excelReader.readXlsxOnClassPath(COUPANG_DICT_URI)) {

            XSSFSheet sheet = wb.getSheetAt(0); // 해당 엑셀파일의 시트수
            int rows = sheet.getPhysicalNumberOfRows(); // 해당 시트의 행의 개수

            // 파싱
            for (int rowIdx = 1; rowIdx < rows; rowIdx++) {
                XSSFRow row = sheet.getRow(rowIdx); // 각 행을 읽어온다

                if (row != null) {
                    XSSFCell cell = row.getCell(6);
                    String key = cell.getStringCellValue();
                    cell = row.getCell(8);
                    String value = cell.getStringCellValue();
                    coupangCountMap.put(value, 0);
                }
            }
        } catch (Exception e) {
            log.error(e.getMessage(), e);
            throw new RuntimeException("쿠팡 판매 목록을 읽어오는 중 에러가 발생했습니다.");
        }

        threadLocalCoupangCountMap.set(coupangCountMap);
        return coupangCountMap;
    }

    /**
     * 쿠팡 주문 목록 엑셀 파일을 읽어 파싱 후 CNP에 업로드할 수 있도록 행별로 정리해 리턴한다.
     *
     * @param file
     * @return CnpInputDto list
     */
    private List<CnpInputDto> parseCoupang(File file) {
        List<CnpInputDto> answer = new ArrayList<>();
        List<CoupangColumnDto> coupangLs = new ArrayList<>();
        Map<String, CoupangColumnDto> ansMap = new HashMap<>();
        Map<String, Integer> coupangCountMap = getCoupangCountMap();

        try (XSSFWorkbook workbook = excelReader.readXlsxFile(file)) {
            XSSFSheet sheet = workbook.getSheetAt(0); // 해당 엑셀파일의 시트수
            int rows = sheet.getPhysicalNumberOfRows(); // 해당 시트의 행의 개수
//		System.out.println(rows);

            XSSFRow row = sheet.getRow(0); // 목록행 읽어오기
            int cells = row.getPhysicalNumberOfCells();

            int receiverNameIdx = -1;
            int receiverPhoneIdx = -1;
            int receiverAddressIdx = -1;
            int productNameIdx = -1;
            int etcIdx = -1;
            int idIdx = -1;
            int quantityIdx = -1;

            for (int colIdx = 0; colIdx < cells; colIdx++) {
                XSSFCell cell = row.getCell(colIdx);
                String menu = cell.getStringCellValue();
                if (menu.equals("구매수(수량)")) quantityIdx = colIdx;
                if (menu.equals("묶음배송번호")) idIdx = colIdx;
                if (menu.equals("수취인이름")) receiverNameIdx = colIdx;
                if (menu.equals("수취인전화번호")) receiverPhoneIdx = colIdx;
                if (menu.equals("수취인 주소")) receiverAddressIdx = colIdx;
                if (menu.equals("노출상품명(옵션명)")) productNameIdx = colIdx;
                if (menu.equals("배송메세지")) etcIdx = colIdx;
            }


            // 파싱
            for (int rowIdx = 1; rowIdx < rows; rowIdx++) {
                row = sheet.getRow(rowIdx); // 각 행을 읽어온다
                CoupangColumnDto coupangData = new CoupangColumnDto();
                CnpInputDto cnpInput = new CnpInputDto();

                if (row != null) {
                    String productName = row.getCell(12).getStringCellValue();

                    // 셀에 담겨있는 값을 읽는다.
                    // 묶음 배송 번호
                    String setId = row.getCell(idIdx).getStringCellValue();
                    // 묶음 배송
                    // 맵에 있는지?

                    CoupangColumnDto inValue = ansMap.get(setId);

                    if (inValue != null) {
                        coupangData = inValue;
                        String displayedProductName = coupangData.getDisplayedProductName();

                        String thisRowDisplayedProductName = row.getCell(productNameIdx).getStringCellValue();
                        Integer quantity = Integer.parseInt(row.getCell(quantityIdx).getStringCellValue());

                        // 선택사항
                        if (thisRowDisplayedProductName.contains("선택사항")) {
                            String[] options = thisRowDisplayedProductName.split(",");
                            String option = options[options.length - 1];
                            String ans = option.substring(COUPANG_OPTION_START_IDX, option.length() - 2);
                            if (quantity == 1) coupangData.setDisplayedProductName(displayedProductName + " // " + ans);
                            else
                                coupangData.setDisplayedProductName(displayedProductName + " // " + ans + " (" + quantity + "개)");
                            coupangCountMap.replace(ans, coupangCountMap.get(ans) + quantity);
                        } else {
                            thisRowDisplayedProductName = coupangDict.get(thisRowDisplayedProductName);
                            if (quantity == 1)
                                coupangData.setDisplayedProductName(displayedProductName + " // " + thisRowDisplayedProductName);
                            else
                                coupangData.setDisplayedProductName(displayedProductName + " // " + thisRowDisplayedProductName + " (" + quantity + "개)");
                            coupangCountMap.replace(thisRowDisplayedProductName, coupangCountMap.get(thisRowDisplayedProductName) + quantity);
                        }

                        continue;
                    }
//				coupangData.setShippingNum(setId);

                    // 구매수
                    int quantity = Integer.parseInt(row.getCell(quantityIdx).getStringCellValue());
                    coupangData.setQuantity(quantity);

                    // 수취인이름
                    coupangData.setReceiverName(row.getCell(receiverNameIdx).getStringCellValue());

                    // 수취인전화번호
                    coupangData.setReceiverPhone(processPhone(row.getCell(receiverPhoneIdx).getStringCellValue()));

                    // 수취인  주소
                    coupangData.setReceiverAddress(row.getCell(receiverAddressIdx).getStringCellValue());

                    // 노출상품명(옵션명)

                    String val = coupangDict.get(productName);
                    String ans = "쿠 - ";
                    // 선택사항
                    if (productName.contains("선택사항")) {
                        String[] options = productName.split(",");
                        String option = options[options.length - 1];
                        ans += option.substring(COUPANG_OPTION_START_IDX, option.length() - 2);

                        for (Entry<String, String> entry : coupangDict.entrySet()) {
                            if (productName.equals(val)) {
                                coupangCountMap.replace(entry.getValue(), coupangCountMap.get(entry.getValue()) + quantity);
                                break;
                            }
                        }
                    } else {
                        productName = coupangDict.get(productName);
                        coupangCountMap.replace(productName, coupangCountMap.get(productName) + quantity);
                        ans += productName;
                    }

                    if (quantity == 1) coupangData.setDisplayedProductName(ans);
                    else coupangData.setDisplayedProductName(ans + " (" + quantity + "개)");
                    // 배송메세지
                    coupangData.setDeliveryMessage(row.getCell(etcIdx).getStringCellValue());
                    ansMap.put(setId, coupangData);
                    coupangLs.add(coupangData);
                }
            }


            for (CoupangColumnDto coupangData : coupangLs) {
                CnpInputDto cnpInput = new CnpInputDto();

                cnpInput.setReceiverName(coupangData.getReceiverName());
                cnpInput.setReceiverPhone(coupangData.getReceiverPhone());
                cnpInput.setReceiverAddress(coupangData.getReceiverAddress());
                cnpInput.setOrderContents(coupangData.getDisplayedProductName());
                cnpInput.setRemark(coupangData.getDeliveryMessage());

                answer.add(cnpInput);
            }
        } catch (Exception e) {
            log.error(e.getMessage(), e);
            throw new RuntimeException("쿠팡 파일을 CNP 입력 양식으로 변환하는 과정에서 문제가 발생했습니다..");
        }

        return answer;
    }


    /**
     * 쿠팡 카운트를 계산할 목록을 생성해 리턴한다.
     *
     * @return
     */
    private List<Pair> coupangCountList() {
        Map<String, Integer> coupangCountMap = getCoupangCountMap();

        List<Pair> answer = new ArrayList<>();
        int total = 0;
        for (Entry<String, Integer> entry : coupangCountMap.entrySet()) {
            answer.add(new Pair(entry.getKey(), entry.getValue()));
            total += entry.getValue();
        }
        Collections.sort(answer);

        answer.add(new Pair("합계", total));

        return answer;
    }


    /**
     * ====================================  쿠팡  ========================================
     */


    /**
     * 빈 생성 시점에 네이버 판매목록 엑셀을 읽어 메모리에 로드한다.
     */
    private void readNaverProducts() {
        try (XSSFWorkbook workbook = excelReader.readXlsxOnClassPath(NAVER_DICT_URI)) {
            XSSFSheet sheet = workbook.getSheetAt(0); // 해당 엑셀파일의 시트수
            int rows = sheet.getPhysicalNumberOfRows(); // 해당 시트의 행의 개수

            // 파싱
            for (int rowIdx = 1; rowIdx < rows; rowIdx++) {
                XSSFRow row = sheet.getRow(rowIdx); // 각 행을 읽어온다

                if (row != null) {

                    if (row.getCell(0) == null) {
                        // 옵션 ->
                        XSSFCell cell = row.getCell(2);
                        String key = cell.getStringCellValue();
                        cell = row.getCell(3);
                        String value = cell.getStringCellValue();
                        naverOptionDict.put(key, value);
                    } else {
                        // 상품명 ->
                        XSSFCell cell = row.getCell(0);
                        String key = cell.getStringCellValue();
                        cell = row.getCell(1);
                        String value = cell.getStringCellValue();
                        naverProductDict.put(key, value);
                    }
                }
            }
        } catch (Exception e) {
            log.error(e.getMessage(), e);
            throw new RuntimeException("네이버 판매 목록을 읽어오는 중 에러가 발생했습니다.");
        }
    }

    /**
     * 현재 요청에서 집계된 네이버 판매량 집계 Map을 판매목록 엑셀을 읽어 생성하거나
     * Thread safe한 HttpServletRequest의 attribute에서 가져온다.
     */
    private Map<String, Integer> getNaverCountMap() {
        if (threadLocalNaverCountMap.get() != null) {
            return threadLocalNaverCountMap.get();
        }
        Map<String, Integer> naverCountMap = new HashMap<>();

        try (XSSFWorkbook wb = excelReader.readXlsxOnClassPath(NAVER_DICT_URI)) {
            XSSFSheet sheet = wb.getSheetAt(0); // 해당 엑셀파일의 시트수
            int rows = sheet.getPhysicalNumberOfRows(); // 해당 시트의 행의 개수

            // 파싱
            for (int rowIdx = 1; rowIdx < rows; rowIdx++) {
                XSSFRow row = sheet.getRow(rowIdx); // 각 행을 읽어온다

                if (row != null) {

                    if (row.getCell(0) == null) {
                        // 옵션 ->
                        XSSFCell cell = row.getCell(2);
                        String key = cell.getStringCellValue();
                        cell = row.getCell(3);
                        naverCountMap.put(key, 0);
                    } else {
                        // 상품명 ->
                        XSSFCell cell = row.getCell(0);
                        String key = cell.getStringCellValue();
                        cell = row.getCell(1);
                        naverCountMap.put(key, 0);
                    }
                }
            }
        } catch (Exception e) {
            log.error(e.getMessage(), e);
            throw new RuntimeException("네이버 판매 목록을 읽어오는 중 에러가 발생했습니다.");
        }

        threadLocalNaverCountMap.set(naverCountMap);
        return naverCountMap;
    }

    /**
     * 네이버 주문 목록 엑셀 파일을 읽어 파싱 후 CNP에 업로드할 수 있도록 행별로 정리해 리턴한다.
     *
     * @param file
     * @return CnpInputDto list
     */
    private List<CnpInputDto> parseNaver(File file) {
        List<CnpInputDto> answer = new ArrayList<>();
        List<NaverColumnDto> naverLs = new ArrayList<>();
        Map<String, NaverColumnDto> map = new HashMap<>();
        Map<String, Integer> naverCountMap = getNaverCountMap();


        try (XSSFWorkbook workbook = excelReader.readXlsxFile(file)) {
            XSSFSheet sheet = workbook.getSheetAt(0); // 해당 엑셀파일의 시트수
            int rows = sheet.getPhysicalNumberOfRows(); // 해당 시트의 행의 개수

            XSSFRow row = sheet.getRow(1); // 목록행 읽어오기
            int cells = row.getPhysicalNumberOfCells();

            int idIdx = -1;
            int receiverNameIdx = -1;
            int receiverPhoneIdx = -1;
            int receiverAddressIdx = -1;
            int receiverAddressDetailIdx = -1;
            int productNameIdx = -1;
            int optionIdx = -1;
            int etcIdx = -1;
            int quantityIdx = -1;

            for (int colIdx = 0; colIdx < cells; colIdx++) {
                XSSFCell cell = row.getCell(colIdx);
                String menu = cell.getStringCellValue();
                if (menu.equals("수량")) quantityIdx = colIdx;
                if (menu.equals("주문번호")) idIdx = colIdx;
                if (menu.equals("수취인명")) receiverNameIdx = colIdx;
                if (menu.equals("수취인연락처1")) receiverPhoneIdx = colIdx;
                if (menu.equals("기본배송지")) receiverAddressIdx = colIdx;
                if (menu.equals("상세배송지")) receiverAddressDetailIdx = colIdx;
                if (menu.equals("상품명")) productNameIdx = colIdx;
                if (menu.equals("옵션정보")) optionIdx = colIdx;
                if (menu.equals("배송메세지")) etcIdx = colIdx;
            }


            // 파싱
            for (int rowIdx = 2; rowIdx < rows; rowIdx++) {
                row = sheet.getRow(rowIdx); // 각 행을 읽어온다
                NaverColumnDto naverData = new NaverColumnDto();
                CnpInputDto cnpInput = new CnpInputDto();

                if (row != null) {
                    // 셀에 담겨있는 값을 읽는다.
                    // 수취인명
                    String receiverName = row.getCell(receiverNameIdx).getStringCellValue();
                    // 수취인 연락처
                    String receiverPhone = processPhone(row.getCell(receiverPhoneIdx).getStringCellValue());
                    String orderNum = row.getCell(idIdx).getStringCellValue();

                    NaverColumnDto chkDto = map.get(orderNum);

                    if (chkDto != null) {
                        // 묶음 배송
                        naverData = chkDto;
                        String displayedProductName = naverData.getProductName();
                        String thisRowDisplayedProductName = row.getCell(productNameIdx).getStringCellValue();
                        String thisRowProcessedProductName = naverProductDict.get(thisRowDisplayedProductName);
                        Integer quantity = (int) row.getCell(quantityIdx).getNumericCellValue();

                        if (row.getCell(optionIdx) != null) {
                            // 옵션이 존재하는 경우
                            String thisRowOption = row.getCell(optionIdx).getStringCellValue();
                            String processedOption = naverOptionDict.get(thisRowOption);

                            if (processedOption != null) {
                                thisRowDisplayedProductName = processedOption;
                                naverCountMap.replace(processedOption, naverCountMap.get(processedOption) + quantity);
                            } else {
                                // 옵션 존재 + naverDict에 등록되지 않은 기존 정책따라가는 상품
                                if (thisRowOption.contains("선택사항"))
                                    thisRowDisplayedProductName = thisRowOption.substring(6, thisRowOption.length());
                                else thisRowDisplayedProductName += " " + thisRowOption;
                            }
                        } else if (thisRowProcessedProductName != null) {
                            // 옵션이 존재하지 않으면서 naverProductDict에 등록된 제품
                            thisRowDisplayedProductName = thisRowProcessedProductName;
                            naverCountMap.replace(thisRowProcessedProductName, naverCountMap.get(thisRowProcessedProductName) + quantity);
                        }


                        if (quantity == 1)
                            naverData.setProductName(displayedProductName + " // " + thisRowDisplayedProductName);
                        else
                            naverData.setProductName(displayedProductName + " // " + thisRowDisplayedProductName + " (" + quantity + "개)");

                        continue;
                    }

                    // 수취인이름
                    naverData.setReceiverName(receiverName);

                    // 수취인연락처1
                    naverData.setReceiverPhone1(receiverPhone);

                    // 구매수
                    int quantity = (int) row.getCell(quantityIdx).getNumericCellValue();
                    naverData.setQuantity(quantity);

                    // 수취인  주소
                    naverData.setReceiverAddress(row.getCell(receiverAddressIdx).getStringCellValue());

                    naverData.setReceiverAddressDetail(row.getCell(receiverAddressDetailIdx).getStringCellValue());

                    // 상품명
                    String itemName = row.getCell(productNameIdx).getStringCellValue();
                    String tag = "네 - ";

                    // 옵션이 없다면
                    if (row.getCell(optionIdx) == null) {
                        // dict 들어갈 부분
                        String processedItemName = naverProductDict.get(itemName);

                        // dict에 존재하는 이름이면
                        if (processedItemName != null) {
                            naverData.setProductName(tag + processedItemName);
                            naverCountMap.replace(processedItemName, naverCountMap.get(processedItemName) + quantity);
                        } else naverData.setProductName(tag + itemName);
                    } else {
                        // 옵션이 있다면
                        String option = row.getCell(optionIdx).getStringCellValue();
                        String processedOptionName = naverOptionDict.get(option);

                        if (processedOptionName != null) {
                            naverData.setProductName(tag + processedOptionName);
                            naverCountMap.replace(processedOptionName, naverCountMap.get(processedOptionName) + quantity);
                        } else {
                            // dict에 없으면 기존 정책에 따라
                            if (option.contains("선택사항"))
                                naverData.setProductName(tag + option.substring(6, option.length()));
                            else naverData.setProductName(tag + itemName + " " + option);
                        }
                    }

                    if (quantity > 1) naverData.setProductName(naverData.getProductName() + " (" + quantity + "개)");
                    // 배송메세지
                    String deliveryMessage = "";
                    if (row.getCell(etcIdx) != null) deliveryMessage = row.getCell(etcIdx).getStringCellValue();
                    naverData.setDeliveryMessage(deliveryMessage);

                    map.put(orderNum, naverData);
                    naverLs.add(naverData);
                }
            }

            for (NaverColumnDto naverData : naverLs) {
                CnpInputDto cnpInput = new CnpInputDto();
                cnpInput.setReceiverName(naverData.getReceiverName());
                cnpInput.setReceiverPhone(naverData.getReceiverPhone1());
                cnpInput.setReceiverAddress(naverData.getReceiverAddress() + " " + naverData.getReceiverAddressDetail());
                cnpInput.setOrderContents(naverData.getProductName());
                cnpInput.setRemark(naverData.getDeliveryMessage());

                answer.add(cnpInput);
            }
        } catch (Exception e) {
            log.error(e.getMessage(), e);
            throw new RuntimeException("네이버 엑셀 파일을 읽어 CNP 양식으로 변환하는 도중 에러가 발생했습니다.");
        }


        return answer;
    }

    /**
     * 네이버 카운트를 계산할 목록을 생성해 리턴한다.
     *
     * @return
     */
    public List<Pair> naverCountList() {
        List<Pair> answer = new ArrayList<>();
        Map<String, Integer> naverCountMap = getCoupangCountMap();
        int total = 0;
        for (Entry<String, Integer> entry : naverCountMap.entrySet()) {
            answer.add(new Pair(entry.getKey(), entry.getValue()));
            total += entry.getValue();
        }
        Collections.sort(answer);

        answer.add(new Pair("합계", total));

        return answer;
    }


    /**
     * ===================================== 운송장 번호 입력용 서비스 메소드 ===============================
     */

    /**
     * 운송장 번호 매핑을 위한 키를 생성
     * ${수취인명_수취인번호} 형태로 생성
     * @param receiverName
     * @param phone
     * @return
     */
    private String getTrackingNumberKey(String receiverName, String phone) {
        return receiverName + "_" + phone.replace("-", "");
    }

    /**
     * 운송장 번호 입력을 위해 주문서 엑셀 파일 읽어 매핑
     * @param file
     * @param shopCode
     * @return
     */
    private Map<String, String> mappingOrderExcelToEnterTrackingNumber(File file, ShopCode shopCode) {
        Map<String, String> map = new HashMap<>();
        int keyColumnIndex = shopCode.getKeyColumnIndex();
        int receiverNameColumnIndex = -1;
        int receiverPhoneColumnIndex = -1;


        try (XSSFWorkbook workbook = excelReader.readXlsxFile(file)) {
            XSSFSheet sheet = workbook.getSheetAt(0); // 해당 엑셀파일의 시트수
            Row firstRow = sheet.getRow(0);

            for (Cell headerCell : firstRow) {
                if (shopCode == ShopCode.COUPANG) {
                    if (excelReader.getCellValueAsString(headerCell).equals("수취인이름")) {
                        receiverNameColumnIndex = headerCell.getColumnIndex();
                    }
                    if (excelReader.getCellValueAsString(headerCell).equals("수취인전화번호")) {
                        receiverPhoneColumnIndex = headerCell.getColumnIndex();
                    }
                } else if (shopCode == ShopCode.NAVER) {
                    if (excelReader.getCellValueAsString(headerCell).equals("수취인명")) {
                        receiverNameColumnIndex = headerCell.getColumnIndex();
                    }
                    if (excelReader.getCellValueAsString(headerCell).equals("수취인연락처1")) {
                        receiverPhoneColumnIndex = headerCell.getColumnIndex();
                    }
                }
            }

            for (Row row : sheet) {
                if (row.getRowNum() == 0) {
                    continue;
                }

                map.put(excelReader.getCellValueAsString(row.getCell(keyColumnIndex))
                        , getTrackingNumberKey(
                                excelReader.getCellValueAsString(row.getCell(receiverNameColumnIndex))
                                , excelReader.getCellValueAsString(row.getCell(receiverPhoneColumnIndex))
                        )
                );
            }
        } catch (Exception e) {
            log.error(e.getMessage(), e);
            throw new RuntimeException("주문서 파일을 읽는 도중 문제가 발생했습니다.");
        }

        return map;
    }

    /**
     * 운송장 번호 입력을 위해 CNPlus에서 다운로드 받은 운송장 번호 엑셀을 읽고 키 : 운송장 번호 매핑한 맵을 리턴
     * @param file
     * @return
     */
    private Map<String, String> readCnpTrackingNumberExcelFile(File file) {
        Map<String, String> answer = new HashMap<>();

        try (XSSFWorkbook workbook = excelReader.readXlsxFile(file)) {
            XSSFSheet sheet = workbook.getSheetAt(0); // 해당 엑셀파일의 시트수
            Row headerRow = sheet.getRow(0);

            int trackingNumberColumnIndex = -1;
            int receiverNameColumnIndex = -1;
            int receiverPhoneNumberColumnIndex = -1;


            for (Cell headerCell : headerRow) {
                if (excelReader.getCellValueAsString(headerCell).equals("운송장번호")) {
                    trackingNumberColumnIndex = headerCell.getColumnIndex();
                }
                if (excelReader.getCellValueAsString(headerCell).equals("받는분")) {
                    receiverNameColumnIndex = headerCell.getColumnIndex();
                }
                if (excelReader.getCellValueAsString(headerCell).equals("받는분전화번호")) {
                    receiverPhoneNumberColumnIndex = headerCell.getColumnIndex();
                }
            }

            for (Row row : sheet) {
                if (row.getRowNum() == 0) {
                    continue;
                }

                answer.put(getTrackingNumberKey(
                                excelReader.getCellValueAsString(row.getCell(receiverNameColumnIndex)),
                                excelReader.getCellValueAsString(row.getCell(receiverPhoneNumberColumnIndex))
                        )
                        , excelReader.getCellValueAsString(row.getCell(trackingNumberColumnIndex)));
            }
        } catch (Exception e) {
            log.error(e.getMessage(), e);
            throw new RuntimeException("CNPlus에서 내려받은 운송장 번호 엑셀 파일을 읽는 중 에러가 발생했습니다.");
        }

        return answer;
    }


    private void enterTrackingNumberOnOrderExcel(File orderExcel, Map<String, String> orderMapping, Map<String, String> trackingNumberMapping, ShopCode shopCode) {
//        String tempFileName = String.format("%s-%s 운송장 업로드.xlsx", LocalDateTime.now().format(formatter), shopCode.getValue());

        int headerIndex = shopCode == ShopCode.COUPANG ? 0 : shopCode == ShopCode.NAVER ? 1 : 0;

        try (XSSFWorkbook xlsxWb = excelReader.readXlsxFile(orderExcel);
             FileOutputStream fos = new FileOutputStream(orderExcel)) {

            XSSFSheet sheet = xlsxWb.getSheetAt(0);
            XSSFRow headerRow = sheet.getRow(headerIndex);

            int trackingNumberColumnIndex = -1;

            for (Cell cell : headerRow) {
                String cellValue = excelReader.getCellValueAsString(cell);

                if (shopCode == ShopCode.COUPANG) {
                    if (cellValue.equals("운송장번호")) {
                        trackingNumberColumnIndex = cell.getColumnIndex();
                        break;
                    }
                }
                else if (shopCode == ShopCode.NAVER) {
                    if (cellValue.equals("송장번호")) {
                        trackingNumberColumnIndex = cell.getColumnIndex();
                        break;
                    }
                }
            }

            if (trackingNumberColumnIndex == -1) {
                log.error("운송장번호 컬럼을 찾을 수 없습니다.");
                throw new RuntimeException("운송장번호 컬럼을 찾을 수 없습니다.");
            }

            for (Row row : sheet) {
                if (row.getRowNum() == headerIndex) {
                    continue;
                }

                String orderDeliveryKey = excelReader.getCellValueAsString(row.getCell(shopCode.getKeyColumnIndex()));
                String trackingNumberKey = orderMapping.get(orderDeliveryKey);
                String trackingNumber = trackingNumberMapping.get(trackingNumberKey);

                if (StringUtils.hasText(trackingNumber)) {
                    row.getCell(trackingNumberColumnIndex).setCellValue(trackingNumber);
                }
            }

            xlsxWb.write(fos);
        } catch (Exception e) {
            log.error(e.getMessage(), e);
            throw new RuntimeException("운송장 번호 입력을 위해 주문 엑셀을 읽어오는 도중 에러가 발생했습니다.");
        }
    }


    /**
     * 핸드폰 번호 양식을 form에 맞게 변경해 리턴한다.
     *
     * @param num
     * @return
     */
    private String processPhone(String num) {
        String answer = "";

        for (int i = 0; i < num.length(); i++) {
            if (num.charAt(i) >= '0' && num.charAt(i) <= '9') answer += num.charAt(i);
        }

        return answer;
    }


    /**
     * 주문 내역 엑셀과 운송장 번호 엑셀을 MultipartFile 파라미터로 받아
     * 운송장 번호 cell 부분을 채워서 다운로드 처리한다.
     *
     * @param orderExcelMultiFile
     * @param trackingNumberMultipartFile
     * @param shopCode
     */
    public void enterTrackingNumber(MultipartFile orderExcelMultiFile, MultipartFile trackingNumberMultipartFile, ShopCode shopCode) {
        File orderExcelTempFile = null;
        File trackingNumberExcelTempFile = null;
        try {
            orderExcelTempFile = fileManager.createFileWithMultipartFile(orderExcelMultiFile);
            trackingNumberExcelTempFile = fileManager.createFileWithMultipartFile(trackingNumberMultipartFile);

            Map<String, String> orderMapping = mappingOrderExcelToEnterTrackingNumber(orderExcelTempFile, shopCode);
            Map<String, String> trackingNumberMapping = readCnpTrackingNumberExcelFile(trackingNumberExcelTempFile);

            enterTrackingNumberOnOrderExcel(orderExcelTempFile, orderMapping, trackingNumberMapping, shopCode);

            fileManager.download(
                    shopCode == ShopCode.COUPANG ? DownloadCode.COUPANG_TRACKING_NUMBER_INPUT
                            : shopCode == ShopCode.NAVER ? DownloadCode.NAVER_TRACKING_NUMBER_INPUT
                            : null,
                    orderExcelTempFile);
        } finally {
            orderExcelTempFile.delete();
            trackingNumberExcelTempFile.delete();
        }
    }
}
