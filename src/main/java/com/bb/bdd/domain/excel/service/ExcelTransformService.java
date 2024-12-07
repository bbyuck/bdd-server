package com.bb.bdd.domain.excel.service;

import com.bb.bdd.domain.excel.DownloadCode;
import com.bb.bdd.domain.excel.ShopCode;
import com.bb.bdd.domain.excel.dto.*;
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
import org.springframework.web.multipart.MultipartFile;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.time.LocalDate;
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

    private final String COUPANG_COUNT_MAP_KEY = "coupang-count";
    private final String NAVER_COUNT_MAP_KEY = "naver-count";
    // 쿠팡 제목 바꾸는 엑셀파일
    private final String COUPANG_DICT_URI = "dict/coupang_dict.xlsx";
    private final String NAVER_DICT_URI = "dict/naver_dict.xlsx";
    private final Integer COUPANG_OPTION_START_IDX = 10;

    private final String COURIER = "대한통운";

    private final ExcelReader excelReader;
    private final FileManager fileManager;

    private static final ThreadLocal<Map<String, Integer>> threadLocalCoupangCountMap = ThreadLocal.withInitial(() -> null);
    private static final ThreadLocal<Map<String, Integer>> threadLocalNaverCountMap = ThreadLocal.withInitial(() -> null);

    private void test() {
        for (Entry<String, String> entry : coupangDict.entrySet()) {
            System.out.println(entry.getKey() + " " + entry.getValue());
        }
    }

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
        String tempFileName = LocalDate.now() + "_" + shopCode.getValue() + "- Cnp.xls";

        // cnp input list
        List<CnpInputDto> cnpInputLs =
                shopCode == ShopCode.COUPANG
                        ? readCoupang(excelFile)
                        : shopCode == ShopCode.NAVER
                        ? readNaver(excelFile) : null;

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
        } catch (IOException e) {
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
            String tempFileName = LocalDate.now() + "_" + shopCode.getValue() + " 판매량.xlsx";

            // sheet 생성
            XSSFSheet sheet = xlsxWb.createSheet(LocalDate.now() + " 판매량");

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
        } catch (IOException e) {
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
        } catch (IOException e) {
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
    private List<CnpInputDto> readCoupang(File file) {
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
        } catch (IOException e) {
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
        } catch (IOException e) {
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
        } catch (IOException e) {
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
    private List<CnpInputDto> readNaver(File file) {
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
                cnpInput.setReceiverAddress(naverData.getReceiverAddress());
                cnpInput.setOrderContents(naverData.getProductName());
                cnpInput.setRemark(naverData.getDeliveryMessage());

                answer.add(cnpInput);
            }
        } catch (IOException e) {
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
     * 운송장 번호 입력을 위해 쿠팡 주문 엑셀을 읽고 파싱한다.
     * @param file
     * @return
     */
    private List<CoupangColumnDto> readCoupangOrderExcelToEnterTrackingNumber(File file) {
        List<CoupangColumnDto> answer = new ArrayList<>();

        try (XSSFWorkbook workbook = excelReader.readXlsxFile(file)) {
            XSSFSheet sheet = workbook.getSheetAt(0); // 해당 엑셀파일의 시트수
            int rows = sheet.getPhysicalNumberOfRows(); // 해당 시트의 행의 개수
//		System.out.println(rows);
            XSSFRow row = sheet.getRow(0);
            int cells = row.getPhysicalNumberOfCells();

            int idxA = -1, idxB = -1, idxC = -1, idxD = -1, idxE = -1, idxF = -1, idxG = -1, idxH = -1, idxI = -1,
                    idxJ = -1, idxK = -1, idxL = -1, idxM = -1,
                    idxN = -1, idxO = -1, idxP = -1, idxQ = -1,
                    idxR = -1, idxS = -1, idxT = -1, idxU = -1,
                    idxV = -1, idxW = -1, idxX = -1, idxY = -1,
                    idxZ = -1, idxAA = -1,
                    idxAB = -1, idxAC = -1, idxAD = -1, idxAE = -1,
                    idxAF = -1, idxAG = -1, idxAH = -1, idxAI = -1,
                    idxAJ = -1, idxAK = -1, idxAL = -1, idxAM = -1, idxAN = -1;

            for (int colIdx = 0; colIdx < cells; colIdx++) {
                String menu = row.getCell(colIdx).getStringCellValue();

                if (menu.equals("번호")) idxA = colIdx;
                if (menu.equals("묶음배송번호")) idxB = colIdx;
                if (menu.equals("주문번호")) idxC = colIdx;
                if (menu.equals("택배사")) idxD = colIdx;
                if (menu.equals("운송장번호")) idxE = colIdx;
                if (menu.equals("분리배송 Y/N")) idxF = colIdx;
                if (menu.equals("분리배송 출고예정일")) idxG = colIdx;
                if (menu.equals("주문시 출고예정일")) idxH = colIdx;
                if (menu.equals("출고일(발송일)")) idxI = colIdx;
                if (menu.equals("주문일")) idxJ = colIdx;
                if (menu.equals("등록상품명")) idxK = colIdx;
                if (menu.equals("등록옵션명")) idxL = colIdx;
                if (menu.equals("노출상품명(옵션명)")) idxM = colIdx;
                if (menu.equals("노출상품ID")) idxN = colIdx;
                if (menu.equals("옵션ID")) idxO = colIdx;
                if (menu.equals("최초등록옵션명")) idxP = colIdx;
                if (menu.equals("업체상품코드")) idxQ = colIdx;
                if (menu.equals("바코드")) idxR = colIdx;
                if (menu.equals("결제액")) idxS = colIdx;
                if (menu.equals("배송비구분")) idxT = colIdx;
                if (menu.equals("배송비")) idxU = colIdx;
                if (menu.equals("도서산간 추가배송비")) idxV = colIdx;
                if (menu.equals("구매수(수량)")) idxW = colIdx;
                if (menu.equals("옵션판매가(판매단가)")) idxX = colIdx;
                if (menu.equals("구매자")) idxY = colIdx;
                if (menu.equals("구매자전화번호")) idxZ = colIdx;
                if (menu.equals("수취인이름")) idxAA = colIdx;
                if (menu.equals("수취인전화번호")) idxAB = colIdx;
                if (menu.equals("우편번호")) idxAC = colIdx;
                if (menu.equals("수취인 주소")) idxAD = colIdx;
                if (menu.equals("배송메세지")) idxAE = colIdx;
                if (menu.equals("상품별 추가메시지")) idxAF = colIdx;
                if (menu.equals("주문자 추가메시지")) idxAG = colIdx;
                if (menu.equals("배송완료일")) idxAH = colIdx;
                if (menu.equals("구매확정일자")) idxAI = colIdx;
                if (menu.equals("개인통관번호(PCCC)")) idxAJ = colIdx;
                if (menu.equals("통관용구매자전화번호")) idxAK = colIdx;
                if (menu.equals("기타")) idxAL = colIdx;
                if (menu.equals("결제위치")) idxAM = colIdx;

            }


            // 파싱
            for (int rowIdx = 1; rowIdx < rows; rowIdx++) {
                row = sheet.getRow(rowIdx); // 각 행을 읽어온다
                CoupangColumnDto coupangData = new CoupangColumnDto();

                if (row != null) {

                    int thisNum = 0;
                    String shippingNum = "";
                    String orderNum = "";
                    String courier = "";
                    String waybillNum = "";
                    String separateDelivery = "";
                    String separateExpectedDeliveryDate = "";
                    String expectedDeliveryDate = "";
                    String deliveryDate = "";
                    String orderDate = "";
                    String productName = "";
                    String optionName = "";
                    String displayedProductName = "";
                    String displayedProductId = "";
                    String optionId = "";
                    String firstOptionName = "";
                    String productCode = "";
                    String barcode = "";
                    Integer payment = null;
                    String deliveryFeeFlag = "";
                    Integer deliveryFee = null;
                    Integer additionalDeliveryFee = null;
                    Integer quantity = null;
                    Integer unitPrice = null;
                    String customerName = "";
                    String customerEmail = "";
                    String customerPhone = "";
                    String receiverName = "";
                    String receiverPhone = "";
                    String postNum = "";
                    String receiverAddress = "";
                    String deliveryMessage = "";
                    String additionalMessagePerItem = "";
                    String ordererAdditionalMessage = "";
                    String deliveryCompleteDate = "";
                    String confirmationPurchaseDate = "";
                    String pccc = "";
                    String buyerPhoneNumForCustomsClearance = "";
                    String etc = "";
                    String paymentLocation = "";

                    if (row.getCell(idxA) != null) thisNum = Integer.parseInt(row.getCell(idxA).getStringCellValue());
                    else throw new IOException();
                    if (row.getCell(idxB) != null) shippingNum = row.getCell(idxB).getStringCellValue();
                    else throw new IOException();
                    if (row.getCell(idxC) != null) orderNum = row.getCell(idxC).getStringCellValue();
                    else throw new IOException();

                    if (row.getCell(idxD) != null) courier = row.getCell(idxD).getStringCellValue();
                    if (row.getCell(idxE) != null) waybillNum = row.getCell(idxE).getStringCellValue();
                    if (row.getCell(idxF) != null) separateDelivery = row.getCell(idxF).getStringCellValue();
                    if (row.getCell(idxG) != null)
                        separateExpectedDeliveryDate = row.getCell(idxG).getStringCellValue();
                    if (row.getCell(idxH) != null) expectedDeliveryDate = row.getCell(idxH).getStringCellValue();
                    if (row.getCell(idxI) != null) deliveryDate = row.getCell(idxI).getStringCellValue();
                    if (row.getCell(idxJ) != null) orderDate = row.getCell(idxJ).getStringCellValue();
                    if (row.getCell(idxK) != null) productName = row.getCell(idxK).getStringCellValue();
                    if (row.getCell(idxL) != null) optionName = row.getCell(idxL).getStringCellValue();
                    if (row.getCell(idxM) != null) displayedProductName = row.getCell(idxM).getStringCellValue();
                    if (row.getCell(idxN) != null) displayedProductId = row.getCell(idxN).getStringCellValue();
                    if (row.getCell(idxO) != null) optionId = row.getCell(idxO).getStringCellValue();
                    if (row.getCell(idxP) != null) firstOptionName = row.getCell(idxP).getStringCellValue();
                    if (row.getCell(idxQ) != null) productCode = row.getCell(idxQ).getStringCellValue();
                    if (row.getCell(idxR) != null) barcode = row.getCell(idxR).getStringCellValue();

                    if (row.getCell(idxS) != null) payment = Integer.parseInt(row.getCell(idxS).getStringCellValue());
                    else throw new IOException();
                    if (row.getCell(idxT) != null) deliveryFeeFlag = row.getCell(idxT).getStringCellValue();
                    else throw new IOException();
                    if (row.getCell(idxU) != null)
                        deliveryFee = Integer.parseInt(row.getCell(idxU).getStringCellValue());
                    else throw new IOException();
                    if (row.getCell(idxV) != null)
                        additionalDeliveryFee = Integer.parseInt(row.getCell(idxV).getStringCellValue());
                    else throw new IOException();
                    if (row.getCell(idxW) != null) quantity = Integer.parseInt(row.getCell(idxW).getStringCellValue());
                    else throw new IOException();
                    if (row.getCell(idxX) != null) unitPrice = Integer.parseInt(row.getCell(idxX).getStringCellValue());
                    else throw new IOException();

                    if (row.getCell(idxY) != null) customerName = row.getCell(idxY).getStringCellValue();
                    if (row.getCell(idxZ) != null) customerPhone = row.getCell(idxZ).getStringCellValue();
                    else throw new IOException();

                    if (row.getCell(idxAA) != null) receiverName = row.getCell(idxAA).getStringCellValue();
                    else throw new IOException();

                    if (row.getCell(idxAB) != null) receiverPhone = row.getCell(idxAB).getStringCellValue();

                    if (row.getCell(idxAC) != null) postNum = row.getCell(idxAC).getStringCellValue();
                    else throw new IOException();

                    if (row.getCell(idxAD) != null) receiverAddress = row.getCell(idxAD).getStringCellValue();
                    if (row.getCell(idxAE) != null) deliveryMessage = row.getCell(idxAE).getStringCellValue();
                    if (row.getCell(idxAF) != null) additionalMessagePerItem = row.getCell(idxAF).getStringCellValue();
                    if (row.getCell(idxAG) != null) ordererAdditionalMessage = row.getCell(idxAG).getStringCellValue();
                    if (row.getCell(idxAH) != null) deliveryCompleteDate = row.getCell(idxAH).getStringCellValue();
                    if (row.getCell(idxAI) != null) confirmationPurchaseDate = row.getCell(idxAI).getStringCellValue();
                    if (row.getCell(idxAJ) != null) pccc = row.getCell(idxAJ).getStringCellValue();
                    if (row.getCell(idxAK) != null)
                        buyerPhoneNumForCustomsClearance = row.getCell(idxAK).getStringCellValue();
                    if (row.getCell(idxAL) != null) etc = row.getCell(idxAL).getStringCellValue();
                    if (row.getCell(idxAM) != null) paymentLocation = row.getCell(idxAM).getStringCellValue();

                    coupangData.setNum(thisNum);
                    coupangData.setShippingNum(shippingNum);
                    coupangData.setOrderNum(orderNum);
                    coupangData.setCourier(courier);
                    coupangData.setWaybillNum(waybillNum);
                    coupangData.setSeparateDelivery(separateDelivery);
                    coupangData.setSeparateExpectedDeliveryDate(separateExpectedDeliveryDate);
                    coupangData.setExpectedDeliveryDate(expectedDeliveryDate);
                    coupangData.setDeliveryDate(deliveryDate);
                    coupangData.setOrderDate(orderDate);
                    coupangData.setProductName(displayedProductName);
                    coupangData.setOptionName(firstOptionName);
                    if (quantity == 1) coupangData.setDisplayedProductName(displayedProductName);
                    else coupangData.setDisplayedProductName(displayedProductName + " (" + quantity + "개)");
                    coupangData.setDisplayedProductId(displayedProductId);
                    coupangData.setOptionId(optionId);
                    coupangData.setFirstOptionName(firstOptionName);
                    coupangData.setProductCode(productCode);
                    coupangData.setBarcode(barcode);
                    coupangData.setPayment(payment);
                    coupangData.setDeliveryFeeFlag(deliveryFeeFlag);
                    coupangData.setDeliveryFee(deliveryFee);
                    coupangData.setAdditionalDeliveryFee(additionalDeliveryFee);
                    coupangData.setQuantity(quantity);
                    coupangData.setUnitPrice(unitPrice);
                    coupangData.setCustomerName(customerName);
                    coupangData.setCustomerEmail(customerEmail);
                    coupangData.setCustomerPhone(customerPhone);
                    coupangData.setReceiverName(receiverName);
                    coupangData.setReceiverPhone(processPhone(receiverPhone));
                    coupangData.setPostNum(postNum);
                    coupangData.setReceiverAddress(receiverAddress);
                    coupangData.setDeliveryMessage(deliveryMessage);
                    coupangData.setAdditionalMessagePerItem(additionalMessagePerItem);
                    coupangData.setOrdererAdditionalMessage(ordererAdditionalMessage);
                    coupangData.setDeliveryCompleteDate(deliveryCompleteDate);
                    coupangData.setConfirmationPurchaseDate(confirmationPurchaseDate);
                    coupangData.setPccc(pccc);
                    coupangData.setBuyerPhoneNumForCustomsClearance(buyerPhoneNumForCustomsClearance);
                    coupangData.setEtc(etc);
                    coupangData.setPaymentLocation(paymentLocation);

                    answer.add(coupangData);
                }
            }
        } catch (IOException e) {
            log.error(e.getMessage(), e);
            throw new RuntimeException("운송장 번호 입력을 위해 쿠팡 주문 목록을 읽던 도중 에러가 발생했습니다.");
        }

        return answer;
    }

    /**
     * 운송장 번호 입력을 위해 네이버 주문 엑셀을 읽고 파싱한다.
     * @param file
     * @return
     */
    private List<NaverColumnDto> readNaverOrderExcelToEnterTrackingNumber(File file) {
        List<NaverColumnDto> answer = new ArrayList<>();

        try (XSSFWorkbook workbook = excelReader.readXlsxFile(file)) {
            XSSFSheet sheet = workbook.getSheetAt(0); // 해당 엑셀파일의 시트수
            int rows = sheet.getPhysicalNumberOfRows(); // 해당 시트의 행의 개수

            // 파싱
            for (int rowIdx = 2; rowIdx < rows; rowIdx++) {
                XSSFRow row = sheet.getRow(rowIdx); // 각 행을 읽어온다
                NaverColumnDto naverData = new NaverColumnDto();

                if (row != null) {
                    int cells = row.getPhysicalNumberOfCells();

                    String productOrderNum = "";
                    // 주문번호
                    String orderNum = "";
                    // 배송방법
                    String shippingMethod = "";
                    // 택배사
                    String courier = "";
                    // 송장번호
                    String waybillNum = "";
                    // 발송일
                    String shippingDate = "";
                    // 수취인명
                    String receiverName = "";
                    // 상품명
                    String productName = "";
                    // 옵션정보
                    String optionInfo = "";
                    // 수량
                    Integer quantity = null;
                    // 배송비 형태
                    String shippingCostForm = "";
                    // 수취인연락처1
                    String receiverPhone1 = "";
                    // 배송지
                    String receiverAddress = "";
                    // 배송메세지
                    String deliveryMessage = "";
                    // 출고지
                    String shipFrom = "";
                    // 결제수단
                    String methodOfPayment = "";
                    // 수수료 과금구분
                    String feeChargingCategory = "";
                    // 수수료결제방식
                    String feePaymentMethod = "";
                    // 결제수수료
                    Integer paymentFee = null;
                    // 매출연동 수수료
                    Integer salesLinkageFee = null;
                    // 정산예정금액
                    Integer estimatedTotalAmount = null;
                    // 유입경로
                    String channel = "";
                    // 구매자 주민등록번호
                    String buyerSSN = "";
                    // 개인통관고유부호
                    String pccc = "";
                    // 주문일시
                    String orderDateTime = "";
                    // 1년 주문건수
                    Integer numOfOrdersPerYear = null;
                    // 구매자ID
                    String buyerId = "";
                    // 구매자명
                    String buyerName = "";
                    // 결제일
                    String paymentDate = "";
                    // 상품종류
                    String productType = "";
                    // 주문세부상태
                    String orderDetailStatus = "";
                    // 주문상태
                    String orderStatus = "";
                    // 상품번호
                    String itemNum = "";
                    // 배송속성
                    String deliveryProperty = "";
                    // 배송희망일
                    String wantDeliveryDate = "";
                    // (수취인연락처1)
                    String _receiverPhone1 = "";
                    // (수취인연락처2)
                    String _receiverPhone2 = "";
                    // (우편번호)
                    String zipcode = "";
                    // (기본주소)
                    String receiverAddress1 = "";
                    // (상세주소)
                    String receiverAddress2 = "";
                    // (구매자연락처)
                    String buyerPhone = "";
                    SimpleDateFormat sdf = new SimpleDateFormat("yyyy-MM-dd HH:mm");
                    TimeZone tz = TimeZone.getTimeZone("Asia/Seoul");
                    sdf.setTimeZone(tz);

                    if (row.getCell(0) != null) productOrderNum = row.getCell(0).getStringCellValue();
                    if (row.getCell(1) != null) orderNum = row.getCell(1).getStringCellValue();
                    if (row.getCell(2) != null) shippingMethod = row.getCell(2).getStringCellValue();
                    if (row.getCell(3) != null) courier = row.getCell(3).getStringCellValue();
                    if (row.getCell(4) != null) waybillNum = row.getCell(4).getStringCellValue();
                    if (row.getCell(5) != null) shippingDate = sdf.format(row.getCell(5).getDateCellValue());
                    if (row.getCell(6) != null) receiverName = row.getCell(6).getStringCellValue();
                    if (row.getCell(7) != null) productName = row.getCell(7).getStringCellValue();
                    if (row.getCell(8) != null) optionInfo = row.getCell(8).getStringCellValue();
                    if (row.getCell(9) != null) quantity = (int) row.getCell(9).getNumericCellValue();
                    if (row.getCell(10) != null) shippingCostForm = row.getCell(10).getStringCellValue();
                    if (row.getCell(11) != null) receiverPhone1 = row.getCell(11).getStringCellValue();
                    if (row.getCell(12) != null) receiverAddress = row.getCell(12).getStringCellValue();
                    if (row.getCell(13) != null) deliveryMessage = row.getCell(13).getStringCellValue();
                    if (row.getCell(14) != null) shipFrom = row.getCell(14).getStringCellValue();
                    if (row.getCell(15) != null) methodOfPayment = row.getCell(15).getStringCellValue();
                    if (row.getCell(16) != null) feeChargingCategory = row.getCell(16).getStringCellValue();
                    if (row.getCell(17) != null) feePaymentMethod = row.getCell(17).getStringCellValue();
                    if (row.getCell(18) != null) paymentFee = (int) row.getCell(18).getNumericCellValue();
                    if (row.getCell(19) != null) salesLinkageFee = (int) row.getCell(19).getNumericCellValue();
                    if (row.getCell(20) != null) estimatedTotalAmount = (int) row.getCell(20).getNumericCellValue();
                    if (row.getCell(21) != null) channel = row.getCell(21).getStringCellValue();
                    if (row.getCell(22) != null) buyerSSN = row.getCell(22).getStringCellValue();
                    if (row.getCell(23) != null) pccc = row.getCell(23).getStringCellValue();
                    if (row.getCell(24) != null) orderDateTime = sdf.format(row.getCell(24).getDateCellValue());
                    if (row.getCell(25) != null) numOfOrdersPerYear = (int) row.getCell(25).getNumericCellValue();
                    if (row.getCell(26) != null) buyerId = row.getCell(26).getStringCellValue();
                    if (row.getCell(27) != null) buyerName = row.getCell(27).getStringCellValue();
                    if (row.getCell(28) != null) paymentDate = sdf.format(row.getCell(28).getDateCellValue());
                    if (row.getCell(29) != null) productType = row.getCell(29).getStringCellValue();
                    if (row.getCell(30) != null) orderDetailStatus = row.getCell(30).getStringCellValue();
                    if (row.getCell(31) != null) orderStatus = row.getCell(31).getStringCellValue();
                    if (row.getCell(32) != null) itemNum = row.getCell(32).getStringCellValue();
                    if (row.getCell(33) != null) deliveryProperty = row.getCell(33).getStringCellValue();
                    if (row.getCell(34) != null) wantDeliveryDate = row.getCell(34).getStringCellValue();
                    if (row.getCell(35) != null) _receiverPhone1 = row.getCell(35).getStringCellValue();
                    if (row.getCell(36) != null) _receiverPhone2 = row.getCell(36).getStringCellValue();
                    if (row.getCell(37) != null) zipcode = row.getCell(37).getStringCellValue();
                    if (row.getCell(38) != null) receiverAddress1 = row.getCell(38).getStringCellValue();
                    if (row.getCell(39) != null) receiverAddress2 = row.getCell(39).getStringCellValue();
                    if (row.getCell(40) != null) buyerPhone = row.getCell(40).getStringCellValue();

                    receiverPhone1 = processPhone(receiverPhone1);

                    naverData.setProductOrderNum(productOrderNum);
                    naverData.setProductName(productName);
                    naverData.setOrderNum(orderNum);
                    naverData.setShippingMethod(shippingMethod);
                    naverData.setCourier(courier);
                    naverData.setWaybillNum(waybillNum);
                    naverData.setShippingDate(shippingDate);
                    naverData.setReceiverName(receiverName);
                    naverData.setProductName(productName);
                    naverData.setOptionInfo(optionInfo);
                    naverData.setQuantity(quantity);
                    naverData.setShippingCostForm(shippingCostForm);
                    naverData.setReceiverPhone1(receiverPhone1);
                    naverData.setReceiverAddress(receiverAddress);
                    naverData.setDeliveryMessage(deliveryMessage);
                    naverData.setShipFrom(shipFrom);
                    naverData.setMethodOfPayment(methodOfPayment);
                    naverData.setFeeChargingCategory(feeChargingCategory);
                    naverData.setFeePaymentMethod(feePaymentMethod);
                    naverData.setPaymentFee(paymentFee);
                    naverData.setSalesLinkageFee(salesLinkageFee);
                    naverData.setEstimatedTotalAmount(estimatedTotalAmount);
                    naverData.setChannel(channel);
                    naverData.setBuyerSSN(buyerSSN);
                    naverData.setPccc(pccc);
                    naverData.setOrderDateTime(orderDateTime);
                    naverData.setNumOfOrdersPerYear(numOfOrdersPerYear);
                    naverData.setBuyerId(buyerId);
                    naverData.setPaymentDate(paymentDate);
                    naverData.setProductType(productType);
                    naverData.setOrderDetailStatus(orderDetailStatus);
                    naverData.setOrderStatus(orderStatus);
                    naverData.setItemNum(itemNum);
                    naverData.setDeliveryProperty(deliveryProperty);
                    naverData.setWantDeliveryDate(wantDeliveryDate);
                    naverData.set_receiverPhone1(_receiverPhone1);
                    naverData.set_receiverPhone2(_receiverPhone2);
                    naverData.setZipcode(zipcode);
                    naverData.setReceiverAddress1(receiverAddress1);
                    naverData.setReceiverAddress2(receiverAddress2);
                    naverData.setBuyerPhone(buyerPhone);

                    answer.add(naverData);
                }
            }
        } catch (IOException e) {
            log.error(e.getMessage(), e);
            throw new RuntimeException("운송장 번호 입력을 위해 네이버 주문 목록을 읽던 도중 에러가 발생했습니다.");
        }
        return answer;

    }

    /**
     * 운송장 번호 입력을 위해 CNPlus에서 다운로드 받은 운송장 번호 엑셀을 읽고 파싱한다.
     * 수취인명 + 수취인 번호를 키로 함.
     * @param file
     * @return
     */
    private HashMap<String, String> readCnpTrackingNumberExcelFile(File file) {
        HashMap<String, String> answer = new HashMap<>();

        try (HSSFWorkbook workbook = excelReader.readXlsFile(file)) {
            HSSFSheet sheet = workbook.getSheetAt(0); // 해당 엑셀파일의 시트수
            int rows = sheet.getPhysicalNumberOfRows(); // 해당 시트의 행의 개수
//		System.out.println(rows);

            // 파싱
            for (int rowIdx = 1; rowIdx < rows; rowIdx++) {
                HSSFRow row = sheet.getRow(rowIdx); // 각 행을 읽어온다
                CnpOutputDto cnpOutput = new CnpOutputDto();


                if (row != null) {
                    // 셀에 담겨있는 값을 읽는다.
                    String waybillNum = row.getCell(7).getStringCellValue();
                    String receiverName = row.getCell(20).getStringCellValue();
                    String receiverPhone = row.getCell(21).getStringCellValue();

                    answer.put(receiverName + receiverPhone, waybillNum);
                }
            }
        } catch (IOException e) {
            log.error(e.getMessage(), e);
            throw new RuntimeException("CNPlus에서 내려받은 운송장 번호 엑셀 파일을 읽는 중 에러가 발생했습니다.");
        }

        return answer;
    }


    private void enterTrackingNumberOnOrderExcel(File orderExcel, List<String> trackingNumberList, ShopCode shopCode) {
        File xlsxFile = new File(LocalDate.now() + "_쿠팡 운송장 업로드.xlsx");
        int headerIndex = shopCode == ShopCode.COUPANG ? 0 : shopCode == ShopCode.NAVER ? 1 : 0;

        try (XSSFWorkbook xlsxWb = excelReader.readXlsxFile(orderExcel);
             FileOutputStream fos = new FileOutputStream(xlsxFile)) {
            XSSFSheet sheet = xlsxWb.getSheetAt(0);
            XSSFRow headerRow = sheet.getRow(headerIndex);

            int trackingNumberColumnIndex = -1;

            for (Cell cell : headerRow) {
                String cellValue = excelReader.getCellValueAsString(cell);
                if (cellValue.equals("운송장번호")) {
                    trackingNumberColumnIndex = cell.getColumnIndex();
                    break;
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
                Cell trackingNumberBodyCell = row.getCell(trackingNumberColumnIndex);
                trackingNumberBodyCell.setCellValue(trackingNumberList.get(row.getRowNum() - 1));
            }

            xlsxWb.write(fos);
        } catch (IOException e) {
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

            List<String> trackingNumbers = new ArrayList<>();
            if (shopCode == ShopCode.COUPANG) {
                List<CoupangColumnDto> coupangLs = readCoupangOrderExcelToEnterTrackingNumber(orderExcelTempFile);
                HashMap<String, String> customerToWaybill = readCnpTrackingNumberExcelFile(trackingNumberExcelTempFile);

                trackingNumbers = coupangLs.stream().map(coupangData -> customerToWaybill.get(coupangData.getReceiverName() + coupangData.getReceiverPhone())).toList();

            } else if (shopCode == ShopCode.NAVER) {
                List<NaverColumnDto> naverLs = readNaverOrderExcelToEnterTrackingNumber(orderExcelTempFile);
                HashMap<String, String> customerToWaybill = readCnpTrackingNumberExcelFile(trackingNumberExcelTempFile);

                trackingNumbers = naverLs.stream().map(naverData -> customerToWaybill.get(naverData.getReceiverName() + naverData.getReceiverPhone1())).toList();
            }
            enterTrackingNumberOnOrderExcel(orderExcelTempFile, trackingNumbers, shopCode);

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
